#pragma once

// FrmCheckRobot dialog
class FrmCheckRobot : public CDialogEx
{
	DECLARE_DYNAMIC(FrmCheckRobot)

public:
	FrmCheckRobot(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmCheckRobot();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DLG_CHECKBOT };
#endif

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnTimer(UINT_PTR nIDEvent);

public:
	CStatic m_lbl_key;
	CColorEdit m_txt_key[5];
	CButtonST m_btn_reset, m_btn_ok;

	int iCheckpass, iChangetxt;
	int iSecmax, demTimer;

	CString szChNum;

public:

	afx_msg CString _GetRandowText();
	afx_msg void KillfocusEditCode(int nIndex);

	afx_msg void OnBnClickedButtonCodeOk();
	afx_msg void OnBnClickedButtonCodeReset();
	afx_msg void OnBnClickedBtnCancel();

	afx_msg	void OnEnKillfocusEditCodeTxt1();
	afx_msg void OnEnKillfocusEditCodeTxt2();
	afx_msg void OnEnKillfocusEditCodeTxt3();
	afx_msg void OnEnKillfocusEditCodeTxt4();
	afx_msg void OnEnKillfocusEditCodeTxt5();

	afx_msg void OnEnChangeEditCodeTxt1();
	afx_msg void OnEnChangeEditCodeTxt2();
	afx_msg void OnEnChangeEditCodeTxt3();
	afx_msg void OnEnChangeEditCodeTxt4();
	afx_msg void OnEnChangeEditCodeTxt5();

};
