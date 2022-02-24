#pragma once
#include "StdAfx.h"
#include "CReadFileCSV.h"
#include "CDownloadFileSever.h"

class DlgListUpdate : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgListUpdate)

public:
	DlgListUpdate(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgListUpdate();

	// Dialog Data
	enum { IDD = DLG_LIST_UPDATE };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);	
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);

public:
	CColorEdit m_lbl_total;
	CButton m_check_all, m_check_skip;
	CListCtrlEx m_list_cate, m_list_view;

	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CThemeHelperST m_ThemeHelper;
	
	Function *ff;
	Base64Ex *bb;
	CDownloadFileSever *sv;
	CReadFileCSV *rfile;
	CRegistry *reg;

	MenuIcon mnIcon;

	_bstr_t bsConfig;
	_WorksheetPtr psConfig;
	RangePtr prConfig;

	CString szpathDefault;	
	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;

	int colType, colCate;
	int colRGB, colIcon, colMH, colND;
	int iOnOK;

	int iwMH, iwND, iDlgW, iDlgH;
	int iAutoUpdate;

	struct MyStrCSV
	{
		// -- Danh mục CSV ------------------------------------------
		// 0: Lĩnh vực - 1: Số hiệu - 2: Nội dung
		// 3: Loại văn bản - 4: CQ ban hành - 5: Ngày ban hành - 6: Ngày hiệu lực - 7: Trạng thái
		// 8: Hyperlink File gốc - 9: Hyperlink File GXD - 10: Hyperlink Thảo luận
		vector<CString> vSubItem;
		int iEnable;
	};

	vector<MyStrCSV> vecItem;

	struct MyStrLinhvuc
	{
		CString szTenlv;
		COLORREF clrBkg;
	};

	vector<MyStrLinhvuc> vecLinhvuc;

	int tslkq, lanshow;
	int iStopload, ibuocnhay;

	CImageList m_imageList;
	CSysImageList m_sysImageList;

	bool bFSever;

	int nItem, nSubItem;
	int nItemCate, nSubItemCate;

public:
	int _CheckUpdateData(bool blUpdate = true);			// <--- Hàm quan trọng check dữ liệu update
	int _LoadListCate();
	int _LoadListView();

	CString _SetDirDefault();
	COLORREF _GetColorLinhvuc(CString szLinhvuc);
	int _GetCheckItemVisible(CString szLinhvuc, int itype = -1);
	int _GetCountItemsCheck(CString szLinhvuc = L"", int iType = -1);

	void _SetDefault();
	void _SelectAllItems();
	void _DeleteAllListCate();
	void _DeleteAllListView();
	void _XacdinhSheetLienquan();
	void _SetStatusKetqua(int num = 0);
	void _AddItemsListKetqua(int iBegin, int iEnd, int itype = 0, int icheck = 1);	
	void _AutoSothutu(_WorksheetPtr pSheet, int numStart = 1, int nRowStart = 0, int nRowEnd = 0);
	void _GetItemSelect(int icheck);
	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedCheckAll();

	afx_msg void OnLsupCheck();
	afx_msg void OnLsupUncheck();
	afx_msg void OnLsupCheckall();
	
	afx_msg void OnNMClickListCate(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListCate(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult);

	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);
};
