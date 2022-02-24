#pragma once

// CMultiFileOpenDialog
class CMultiFileOpenDialog : public CFileDialog
{
// Attribbutes
private:
   LPTSTR m_pFileBuff;
   BOOL m_bUseExtBuffer;

// Construction
public:
	CMultiFileOpenDialog(BOOL bVistaStyle,
		LPCTSTR lpszDefExt = NULL,
		LPCTSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_EXPLORER,
		LPCTSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL);
	virtual ~CMultiFileOpenDialog();

// Operations
public:
   void SetUseExtBuffer(BOOL bNewValue) {m_bUseExtBuffer = bNewValue;}

// Overrides
protected:
   virtual void OnFileNameChange();

// Implementation
private:
   DWORD _CalcRequiredBuffSize();
   DWORD _CalcRequiredBuffSizeOldStyle();
   DWORD _CalcRequiredBuffSizeVistaStyle();
   void _SetExtBuffer(DWORD dwReqBuffSize);

// Message handlers
protected:
	DECLARE_MESSAGE_MAP()
   
};


