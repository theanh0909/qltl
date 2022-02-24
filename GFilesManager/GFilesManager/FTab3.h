#pragma once
#include "pch.h"

class FTab3 : public CDialogEx
{
	DECLARE_DYNAMIC(FTab3)

public:
	FTab3(CWnd* pParent = NULL);   // standard constructor
	virtual ~FTab3();

	// Dialog Data
	enum { IDD = FV_TAB3 };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);

};
