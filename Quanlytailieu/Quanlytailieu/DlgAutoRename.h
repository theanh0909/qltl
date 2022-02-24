#pragma once
#include "StdAfx.h"
#include "FrmTabSTT.h"
#include "FrmTabAdd.h"
#include "FrmTabDelete.h"
#include "FrmTabReplace.h"
#include "FrmTabKhac.h"

class DlgAutoRename : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgAutoRename)

public:
	DlgAutoRename(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgAutoRename();

	// Dialog Data
	enum { IDD = DLG_AUTORENAME };

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
	afx_msg void OnTimer(UINT_PTR nIDEvent);

public:
	CTabCtrl m_tab_main;
	CSplitButton m_splbtn_browse;
	CXPStyleButtonST m_btn_ok, m_btn_cancel, m_btn_apply;
	CXPStyleButtonST m_btn_reset, m_btn_undo, m_btn_redo;
	CThemeHelperST m_ThemeHelper;
	CColorEdit m_txt_browse, m_lbl_total, m_lbl_hdan;
	CListCtrlEx m_list_view;
	CImageList m_imageList;
	CSysImageList m_sysImageList;
	MenuIcon mnIcon;

	// nIndexSplit = 0: Danh sách thư mục | 1: Danh sách files
	int nIndexSplit;

	// iTabDefault = 0: Số thứ tự | 1: Thêm | 2: Xóa | 3: Thay thế | 4: Khác
	int iTabDefault;

	int nItem, nSubItem;

	CString szTargetDirectory;

	int iGFLTenNew, iGFLTenOld, iGFLPath;
	int iTypeApply;

	int nTotalCol, iwCol[4];
	int iDlgW, iDlgH;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	struct MyStrItems
	{
		CString szNew;
		CString szName;
		CString szType;
		CString szFullpath;
		int nIndex;
	};

	vector< MyStrItems> vecItems;

	FrmTabSTT* m_frmSTT;
	FrmTabAdd* m_frmAdd;
	FrmTabDelete* m_frmDel;
	FrmTabReplace* m_frmRep;
	FrmTabKhac* m_frmKhac;

	int nIndexVec, nCountList;
	vector<vector<CString>> vecList;

	int iKeyESC;
	CString szBefore, szAfter;

public:
	void _SetDefault();
	void _Timkiemdulieu();
	void _CreateTabMain();
	void _LoadAllKetqua();
	void _SetTotalKetqua();
	void _DeleteListKetqua();	
	void _SelectAllItems();
	void _DeleteAllImages();	
	void _GetItemSelect(int icheck);

	void _SaveRegistry();
	void _LoadRegistry();
	void _LoadRegistryTab();
	void _SetLocationAndSize();
	
	int _TimkiemFiles(CString szPath = L"");
	int _TimkiemThumuc(CString szPath = L"");	
	int _GetAllSelectedItems(vector<int> &vecRow);

	void _RunFuncSothutu();
	void _RunFuncAdd();
	void _RunFuncDelete();
	void _RunFuncReplace();
	void _RunFuncKhac();

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedSplitBrowse();
	afx_msg void OnBnClickedBtnApdung();
	afx_msg void OnBnClickedCheckAll();
	afx_msg void OnBnClickedBtnReset();
	afx_msg void OnBnClickedBtnUndo();
	afx_msg void OnBnClickedBtnRedo();

	afx_msg void OnEnSetfocusEditList();
	afx_msg void OnEnKillfocusEditList();
	afx_msg void OnEnSetfocusTxtBrowse();

	afx_msg void OnAutornFolder();
	afx_msg void OnAutornFiles();

	afx_msg void OnAutornApply();
	afx_msg void OnAutornOpenfile();
	afx_msg void OnAutornOpenfolder();

	afx_msg void OnAutornUp();
	afx_msg void OnAutornDown();
	afx_msg void OnAutornReverse();

	afx_msg void OnAutornChfiles();
	afx_msg void OnAutornChfolder();

	afx_msg void OnAutornSel();
	afx_msg void OnAutornUnsel();
	afx_msg void OnAutornSelall();
	
	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnTcnSelchangeTabmain(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult);
};
