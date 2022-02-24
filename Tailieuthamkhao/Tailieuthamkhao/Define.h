#pragma once

/* Thay đổi tên file sử dụng phần mềm tại đây */

#define ASCO		// DU_TOAN | DU_THAU | HQDA | QUYET_TOAN | QLCL | QLDA | ASCO

/*******************************************************************/

#if defined(DU_TOAN)
	#define TENFILESUDUNG			L"TailieuDutoan"	
	#define LicenseDll				L"dWTURUmhRi=HKTjMTbzlR0Wm1SUU"
#elif defined(DU_THAU)
	#define TENFILESUDUNG			L"TailieuDuthau"
	#define LicenseDll				L"R3NWRkpZ3peOqFOb0IYcSzJxS1pZPW89"
#elif defined(HQDA)
	#define TENFILESUDUNG			L"TailieuHQDA"
	#define LicenseDll				L"NULL"
#elif defined(QUYET_TOAN)
	#define TENFILESUDUNG			L"TailieuQuyettoan"
	#define LicenseDll				L"bXNEbERUkxUi0kSSVGWaGTNJO2JMRVVo"
#elif defined(QLCL)
	#define TENFILESUDUNG			L"TailieuQLCL"
	#define LicenseDll				L"R3NUVmpRXpeVFFSQ0FYUOkdx4lRZPU09"
#elif defined(ASCO)
	#define TENFILESUDUNG			L"TailieuASCO"
	#define LicenseDll				L"R3NFRGBxlpeU5FPV0yWQOldxkmMZPXY9"
#elif defined(QLDA)
	#define TENFILESUDUNG			L"TailieuQLDA"
	#define LicenseDll				L"NULL"
#else
	#define TENFILESUDUNG			L"NULL"
	#define LicenseDll				L"NULL"
#endif

/*******************************************************************/

#define FILENAMEDLL				L"Tailieuthamkhao.dll"
#define ExcelApp				L"Excel.Application"

#define FOLDERSAVE				L"Thuvien"
#define FOLDERDOWNLOAD			L"Download"

#define FILETIEUCHUAN			L"TieuchuanXD"
#define FILELISTDOWNLOAD		L"Listdownload.mdb"
#define PATH_DOWNLOAD			L"DownloadTailieu.txt"
#define ODBC_Driver				L"MICROSOFT ACCESS DRIVER (*.mdb, *.accdb)"

#define THONGBAO				L"Thông báo"
#define FONTNORMAL				L"MS Shell Dlg"

#define SIZE_MENU_ICON			16
#define MINISEC_TIME_OUT		0
#define MAX_LEN					1024

#define CHECKCONN				L"Mm5HOWJiVXVZNVpNnlM06jZV33SSvm9pkFJYdkwzeWRkbjM1"
#define LOCALSEVER				L"51MlpMTFIxYkdsbGRTTUMcWBOO=D1RH2aS5DBkkEETa2RYUnZZV0dlNT1uUW"
#define FILEDOWNLOAD			L"cG01VGJsVgR=DJTV=VPOlwUVFmhhh3hWVFY9Uk1D"

#define FOLDER_SEVER			L"lNT1ggjlVi=T0UjaUWwm91kFFWPW"
#define FILE_CATEGORY			L"bXVVemZcExbh1EUemXUMkTJNGFpbPXZr"
#define FILE_SAVE				L"cG01WmJsVwR=DJTV=VPOkwZRYDRda1pYWVY9ellT"

#define LicenseFunc				L"MjBFRGBxlJWUumPV1YWVClhVSWVdPWw9"
#define CreateKeyRun			L"MFBGTlVYTmxkSFJwYm1kemZcEheF1lUeEGVRVTAJ51ZZa1ZjUjFoRVNsRVhOPURF"