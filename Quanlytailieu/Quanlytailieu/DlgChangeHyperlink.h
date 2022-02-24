#pragma once
#include "StdAfx.h"

class DlgChangeHyperlink : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgChangeHyperlink)

public:
	DlgChangeHyperlink(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgChangeHyperlink();

	// Dialog Data
	enum { IDD = DLG_CHANGE_HYPERLINK };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);

public:
	CColorEdit m_lbl_hdan;
	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CXPStyleButtonST m_btn_add, m_btn_del, m_btn_up, m_btn_down;	
	CThemeHelperST m_ThemeHelper;

	CColorEdit m_txt_dirOld, m_txt_dirNew;
	CListCtrlEx m_list_view;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	CString pathFolderDir;
	CString szDirOld, szDirNew;

	int iColumnPath, iColumnFile;
	int iGFLType, iGFLPathOld, iGFLPathNew;
	int nTotalCol, iwCol[4];
	int iDlgW, iDlgH;

	int nItem, nSubItem;
	MenuIcon mnIcon;

	struct MyStrRep
	{
		CString szOLD;
		CString szNEW;
	};

	vector< MyStrRep> vecREP;

public:
	void _SetDefault();

	void _SetLocationAndSize();
	void _SaveListReplace();
	void _LoadListReplace();
	void _SaveRegistry();
	void _LoadRegistry();
	void _ReplaceHyperlink(int nRow);
	int _GetAllSelectedItems(vector<int> &vecRow);

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();

	afx_msg void OnBnClickedBtnOld();
	afx_msg void OnBnClickedBtnNew();
	afx_msg void OnBnClickedBtnAdd();
	
	afx_msg void OnBnClickedBtnDel();
	afx_msg void OnBnClickedBtnUp();
	afx_msg void OnBnClickedBtnDown();

	afx_msg void OnRepAdd();
	afx_msg void OnRepDel();
	afx_msg void OnRepUp();
	afx_msg void OnRepDown();
	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult);
};
