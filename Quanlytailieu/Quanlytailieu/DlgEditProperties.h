#pragma once
#include "StdAfx.h"

class DlgEditProperties : public CDialogEx
{
	DECLARE_DYNAMIC(DlgEditProperties)

public:
	DlgEditProperties(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgEditProperties();

	// Dialog Data
	enum { IDD = DLG_EDITPROPERTIES };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
	afx_msg void OnMaxMinInfo(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSplitterDelta(NMHDR* pNMHDR, LRESULT* pResult);

public:

	CListCtrlEx m_list_view;
	CColorEdit m_lbl_progress;
	CStatic m_lbl_info, m_lbl_kq;
	CSplitterControl m_wndSplitterVer;	

	CColorEdit m_txt_search, m_txt_filename;
	CColorEdit m_txt_title, m_txt_subject, m_txt_tags, m_txt_categories;
	CColorEdit m_txt_commments, m_txt_authors, m_txt_company;

	CXPStyleButtonST m_btn_search, m_btn_update, m_btn_cancel;
	CXPStyleButtonST m_btn_up, m_btn_down, m_btn_clear, m_btn_saved;
	CThemeHelperST m_ThemeHelper;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int iStopload, lanshow, ibuocnhay;
	long tongkq, tongfilter;

	int iGFLType, iGFLTen, iGFLTitle, iGFLSubject, iGFLTags, iGFLCategories, 
		iGFLComments, iGFLAuthors, iGFLCompany, iGFLPath;
	int nItem, nSubItem;
	int nCheckEdit;
	int iKeyESC;

	int jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType, jcotNoidung, jcotLink;

	int nTotalCol, iwCol[11];
	int iDlgW, iDlgH;

	bool blRead[7];
	bool blIcon;

	CString pathFolderDir;
	CString szBefore, szAfter;
	CString szCopy;

	CImageList m_imageList;
	CSysImageList m_sysImageList;

	MenuIcon mnIcon;

	struct MyStrProperties
	{
		int iID;
		int nRow;
		CString szTen;
		CString szType;
		CString szNoidungEx;

		CString szTitle;
		CString szSubject;
		CString szTags;
		CString szCategories;
		CString szComments;

		CString szAuthors;
		CString szCompany;

		CString szPath;
	};

	vector<MyStrProperties> vecItem;
	vector<MyStrProperties> vecFilter;

public:
	void _GetDanhsachFiles(int iSelectCell = -1);	// <-- Hàm quan trọng xác định số lượng files sử dụng
	void _XacdinhSheetLienquan();
	void _Timkiemdulieu();

	void _SetDefault();	
	void _ResizeDialog();
	void _SetSpliterPane();	
	void _SetBkgColorText(int nIndex, CColorEdit &m_txt_read);
	void _GetGroupRowStartEnd(int &iRowStart, int &iRowEnd);

	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();

	void _AddItemsListKetqua(int iBegin, int iEnd);	
	void _MoveDlgItem(int nD, const CRect& rcPos, BOOL bRepaint);
	void _SetBannerTextMultiLine(CColorEdit &m_txt_multi, bool blSetfocus, CString szBanner);
	void _SetAllBannerMulti(bool blFocus = false);
	void _SetColorCopy(CColorEdit &m_txt_copy);
	void _SetIconButtonSaved(int itype);
	void _LoadProperties(int nIndex);
	void _SetStatusKetqua(int num);
	void _SetSaveDocument(int num);
	void _SaveAllFieldsFix();
	void _SetTextCopy();

	int _FilterData();
	int _GetAllSelectedItems(vector<int> &vecRow);		

	afx_msg void OnBnClickedStt2();
	afx_msg void OnBnClickedStt3();
	afx_msg void OnBnClickedStt4();
	afx_msg void OnBnClickedStt5();
	afx_msg void OnBnClickedStt6();
	afx_msg void OnBnClickedStt7();
	afx_msg void OnBnClickedStt8();

	afx_msg void OnBnClickedBtnUp();
	afx_msg void OnBnClickedBtnDown();
	afx_msg void OnBnClickedBtnClear();
	afx_msg void OnBnClickedBtnSaved();	

	afx_msg void OnBnClickedBtnSearch();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedBtnUpdate();

	afx_msg void OnEnSetfocusTxtSearch();

	afx_msg void OnEnSetfocusEditView();
	afx_msg void OnEnKillfocusEditView();

	afx_msg void OnPropertiesOpenfile();
	afx_msg void OnPropertiesOpenfolder();
	afx_msg void OnPropertiesCopy();
	afx_msg void OnPropertiesPaste();

	afx_msg void OnEnSetfocusTxtTitle();
	afx_msg void OnEnKillfocusTxtTitle();
	afx_msg void OnEnSetfocusTxtSubject();
	afx_msg void OnEnKillfocusTxtSubject();
	afx_msg void OnEnSetfocusTxtTags();
	afx_msg void OnEnKillfocusTxtTags();
	afx_msg void OnEnSetfocusTxtCategories();
	afx_msg void OnEnKillfocusTxtCategories();
	afx_msg void OnEnSetfocusTxtComments();
	afx_msg void OnEnKillfocusTxtComments();

	afx_msg void OnEnChangeTxtTitle();
	afx_msg void OnEnChangeTxtSubject();
	afx_msg void OnEnChangeTxtTags();
	afx_msg void OnEnChangeTxtCategories();
	afx_msg void OnEnChangeTxtComments();
	afx_msg void OnEnChangeTxtAuthors();
	afx_msg void OnEnChangeTxtCompany();

	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);		
};
