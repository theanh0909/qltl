#pragma once
#include "StdAfx.h"

class DlgCopyData : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgCopyData)

public:
	DlgCopyData(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgCopyData();

	// Dialog Data
	enum { IDD = DLG_COPY_DATA };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CButton m_check_all;
	CListCtrlEx m_list_view;

	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CThemeHelperST m_ThemeHelper;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	_bstr_t bsCategory;
	_WorksheetPtr psCategory;
	RangePtr prCategory;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType, jcotNoidung, jcotLink;

	int colLV, colSL;
	int iOnOK;

	int tongCopy;
	int iwLV, iwSL, iDlgW, iDlgH;

	struct MyStrCategory
	{
		CString szLinhvuc;
		CString szKodau;		
		vector<int> irow;
		bool blEnable;
	};

	vector<MyStrCategory> vecItem;

	struct MyStrSheet
	{
		CString szName;
		CString szCode;
		CString szKodau;
		CString szLinhvuc;
	};

	int tslkq, lanshow;
	int iStopload, ibuocnhay;

public:
	int _LoadAllCategory();			// <-- Hàm quan trọng load toàn bộ lĩnh vực
	int _LoadListView();

	void _SetDefault();
	void _DeleteAllList();
	void _XacdinhSheetLienquan();
	void _AddItemsListKetqua(int iBegin, int iEnd, int icheck = 1);
	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();

	void _CopyStyleCategory(_WorksheetPtr pSheet);
	bool _UpdateCategory(_WorksheetPtr pSheet, int nIndex, int &dem);
	void _AutoSothutu(_WorksheetPtr pSheet, int numStart = 1, int nRowStart = 0, int nRowEnd = 0);;

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedCheckAll();

	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);
};
