#include "pch.h"
#include "MyTabCtrl.h"
#include "FTab1.h"
#include "FTab2.h"
#include "FTab3.h"


CMyTabCtrl::CMyTabCtrl()
{
	m_tabPages[0]=new FTab1;
	m_tabPages[1]=new FTab2;
	m_tabPages[2]=new FTab3;
}

CMyTabCtrl::~CMyTabCtrl()
{
	delete m_tabPages[0];
	delete m_tabPages[1];
	delete m_tabPages[2];
}

void CMyTabCtrl::Init(int idefault)
{
	m_tabPages[0]->Create(FV_TAB1, this);
	m_tabPages[1]->Create(FV_TAB2, this);
	m_tabPages[2]->Create(FV_TAB3, this);
	
	m_tabCurrent = idefault;
	switch (m_tabCurrent)
	{
		case 0:
			m_tabPages[0]->ShowWindow(SW_SHOW);
			m_tabPages[1]->ShowWindow(SW_HIDE);
			m_tabPages[2]->ShowWindow(SW_HIDE);
			break;

		case 1:
			m_tabPages[0]->ShowWindow(SW_HIDE);
			m_tabPages[1]->ShowWindow(SW_SHOW);
			m_tabPages[2]->ShowWindow(SW_HIDE);
			break;

		case 2:
			m_tabPages[0]->ShowWindow(SW_HIDE);
			m_tabPages[1]->ShowWindow(SW_HIDE);
			m_tabPages[2]->ShowWindow(SW_SHOW);
			break;

		default:
			break;
	}	

	SetRectangle(m_tabCurrent);
}

void CMyTabCtrl::SetRectangle(int idefault)
{
	CRect tabRect, itemRect;

	GetClientRect(&tabRect);
	GetItemRect(0, &itemRect);

	int nX = itemRect.left;
	int nY = itemRect.bottom+1;
	int nXc = tabRect.right-itemRect.left-1;
	int nYc = tabRect.bottom-nY-1;

	switch (idefault)
	{
		case 0:
			m_tabPages[0]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_SHOWWINDOW);
			m_tabPages[1]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_HIDEWINDOW);
			m_tabPages[2]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_HIDEWINDOW);
			break;

		case 1:
			m_tabPages[0]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_HIDEWINDOW);
			m_tabPages[1]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_SHOWWINDOW);
			m_tabPages[2]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_HIDEWINDOW);
			break;

		case 2:
			m_tabPages[0]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_HIDEWINDOW);
			m_tabPages[1]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_HIDEWINDOW);
			m_tabPages[2]->SetWindowPos(&wndTop, nX, nY, nXc, nYc, SWP_SHOWWINDOW);
			break;

		default:
			break;
	}	
}

BEGIN_MESSAGE_MAP(CMyTabCtrl, CTabCtrl)
	ON_WM_LBUTTONDOWN()
END_MESSAGE_MAP()

void CMyTabCtrl::OnLButtonDown(UINT nFlags, CPoint point) 
{
	CTabCtrl::OnLButtonDown(nFlags, point);
	if(m_tabCurrent != GetCurFocus())
	{
		m_tabPages[m_tabCurrent]->ShowWindow(SW_HIDE);
		m_tabCurrent=GetCurFocus();

		m_tabPages[m_tabCurrent]->ShowWindow(SW_SHOW);
		m_tabPages[m_tabCurrent]->SetFocus();
	}
}
