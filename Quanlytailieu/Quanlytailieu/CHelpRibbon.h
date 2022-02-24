#pragma once
#include "StdAfx.h"

class CHelpRibbon
{

public:
	CHelpRibbon();
	~CHelpRibbon();

	Function *ff;
	CString pathFolder;

public:
	CString _GetLineTextHelp(int nIndex);
	bool _OpenFileHelp(CString szText);

	void _HelpUM();
	void _HelpFAQ();
	void _HelpVideo();
	void _HelpForum();
	void _HelpSoft();
	void _HelpSupport();
	void _HelpFeedback();
};

