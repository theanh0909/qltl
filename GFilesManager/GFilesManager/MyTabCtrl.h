#pragma once
#include "pch.h"

/////////////////////////////////////////////////////////////////////////////
// CMyTabCtrl window

class CMyTabCtrl : public CTabCtrl
{
	// Construction
	public:
		CMyTabCtrl();
		CDialog *m_tabPages[3];
		int m_tabCurrent;

		// Leo edit 26.11.2018
		void Init(int idefault);
		void SetRectangle(int idefault);

	// Implementation
	public:
		virtual ~CMyTabCtrl();

		// Generated message map functions
	protected:
		//{{AFX_MSG(CMyTabCtrl)
		afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
		//}}AFX_MSG

		DECLARE_MESSAGE_MAP()
};
