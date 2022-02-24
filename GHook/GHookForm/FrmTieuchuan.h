#pragma once
#include "CDownloadFileSever.h"

// FrmTieuchuan dialog
class FrmTieuchuan : public CDialogEx
{
	DECLARE_DYNAMIC(FrmTieuchuan)

public:
	FrmTieuchuan(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmTieuchuan();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = FRM_TIEUCHUAN };
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

	int icotID, icotIcon, icotName;
	int wcID, wcIcon, wcTen;
	int iUnikey;

	CString szTukhoa;
	CString szpathDefault;

	bool blChange;

	int demTimer;
	int tslkq, lanshow;
	int iStopload, ibuocnhay;

	struct MyStrTieuchuan
	{
		int nID;
		CString szMa;
		CString szTen;
		CString szPloai;
		CString szGchu;
		CString szDownload;
		CString szTimkiem;
	};

	vector<MyStrTieuchuan> vecTieuchuan;
	vector<MyStrTieuchuan> vecFilter;

	CDownloadFileSever *sv;

	CImageList m_imageList;
	CSysImageList m_sysImageList;

public:
	afx_msg int LoadListView();
	afx_msg int ReadFileTieuchuan();
	afx_msg int LocDulieuTieuchuan();	
	afx_msg int GetSelectItems(vector<int> &vecItem);
	afx_msg CString DownloadFileTieuchuan(CString szFileTchuan);
	afx_msg CString DownloadFileQuytrinh(CString szFileQtrinh);
	afx_msg bool CheckArrowDownNext();
	
	afx_msg void PutKey();
	afx_msg void SetDefault();
	afx_msg void SetLocation();
	afx_msg void DeleteAllList();
	afx_msg void DeleteListIcon();
	afx_msg void ShowResultNext();
	afx_msg void SetTimerSearch();	
	afx_msg void TimkiemTieuchuan();
	afx_msg void ReplaceTukhoa(CString &szKey);
	afx_msg void SelectListItem(bool blEnter = false);
	afx_msg void AddItemsListKetqua(int iBegin, int iEnd);	

	afx_msg void HotroDownloadQuytrinh();

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnEnChangeTxtSearch();

	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnEndScrollListSearch(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListSearch(NMHDR *pNMHDR, LRESULT *pResult);
};
