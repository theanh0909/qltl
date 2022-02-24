#pragma once
#include "pch.h"
#include "CDownloadFileSever.h"

class CRegistryContextMenu
{

public:
	CRegistryContextMenu();
	~CRegistryContextMenu();

public:
	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;
	CDownloadFileSever *sv;

	void _StartRegContextMenu();

};

