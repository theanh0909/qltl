#include "pch.h"

class MyHook
{
	public:
		static MyHook& Instance()
		{
			static MyHook myHook;
			return myHook;
		}	
		HHOOK hook;	
		void InstallHook();
		void UninstallHook();

		MSG msg;
		int Messsages();
};
LRESULT WINAPI MyMouseCallback(int nCode, WPARAM wParam, LPARAM lParam);

