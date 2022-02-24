#pragma once

// FrmQuytrinh dialog
class FrmQuytrinh : public CDialogEx
{
	DECLARE_DYNAMIC(FrmQuytrinh)

public:
	FrmQuytrinh(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmQuytrinh();

	// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = FRM_QUYTRINH };
#endif

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);

public:
	CColorEdit m_txt_search;
	CListCtrlEx m_list_view;

	int icotID, icotQtrinh, icotName;
	int wcID, wcQTrinh, wcTen;
	int iUnikey;

	CString szTukhoa;
	CString szpathDefault;

	bool blChange;

	int demTimer;
	int tslkq, lanshow;
	int iStopload, ibuocnhay;

	int icolSTT, iColMaso, iColNdung, iColChidan;
	int iColURL, iColLienhe, iColGChu, iColType;

	struct MyStrQuytrinh
	{
		int nID;
		CString szMa;
		CString szQuytrinh;
		CString szTenFile;
		CString szTimkiem;
	};

	vector<MyStrQuytrinh> vecQuytrinh;
	vector<MyStrQuytrinh> vecFilter;

public:
	afx_msg int LoadListView();
	afx_msg int ReadFileQuytrinh();
	afx_msg int LocDulieuQuytrinh();
	afx_msg int GetSelectItems(vector<int> &vecItem);
	afx_msg bool CheckArrowDownNext();

	afx_msg void PutKey();
	afx_msg void SetDefault();
	afx_msg void SetLocation();
	afx_msg void DeleteAllList();
	afx_msg void ShowResultNext();
	afx_msg void SetTimerSearch();
	afx_msg void TimkiemQuytrinh();
	afx_msg void CopyQuytrinh(int nID);
	afx_msg void ReplaceTukhoa(CString &szKey);
	afx_msg void SelectListItem(bool blEnter = false);
	afx_msg void AddItemsListKetqua(int iBegin, int iEnd);

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnEnChangeTxtSearch();

	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnEndScrollListSearch(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListSearch(NMHDR *pNMHDR, LRESULT *pResult);
};
