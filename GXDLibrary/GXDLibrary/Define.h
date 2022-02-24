#pragma once

#define VERSION_SOFT			L"1"

#define FILENAMEDLL				L"GXDLibrary.dll"

#define FOLDERDULIEU			L"Dulieu"
#define FOLDERTHUVIEN			L"Thuvien"

#define FileDanhmucvanban		L"Danhmucvanban.mdb"
#define FileLog					L"FileLog.txt"

#define ExcelApp				L"Excel.Application"
#define ODBC_Driver				L"MICROSOFT ACCESS DRIVER (*.mdb, *.accdb)"
#define FontTimes				L"Times New Roman"
#define FontArial				L"Arial"

#define DF_NULL					L"Null"
#define DF_ERROR				L"Error"
#define DF_OPENFILE				L"Mở file"
#define MS_THONGBAO				L"Thông báo"

#define shConfig				L"shConfig"

#define MINISEC_TIME_OUT		0
#define COLW_FIX				20
#define SIZE_MENU_ICON			16
#define SIZE_LINE_CSV			2048
#define ROW_LINE_CSV			8192
#define MAX_LEN					1024
#define MAX_ROW					1000000
#define MAX_RANGE				L"A1:AZ1000000"

#define GXDVN					L"https://gxd.vn"
#define LinkTeamviewGXD			L"https://dl.dropboxusercontent.com/s/garemgss9nfp1ty/Hotro-GXD-Teamviews.exe?dl=0"

#define CHECKCONN				L"JsbXhaTE52ZzRUlVY4=l6RTIYMMUhZSGNTeTJiOT1uOD"
#define SEVERMAIN				L"51MlpMTFIxYkdsbGRTTUMcWBOO=D1RH2aS5DBkkEETa2RYUnZZV0dlNT1uUW"
#define SEVERNEXT				L"k3SDliZWg5bH9IEVVW=EEUnXYaMUcR5GJVMndZRj16PX"
#define SEVERDATA				L"MzRVemZcEVam1GUeFFVeGkdFWmRMdlQ4"

#define shNegs					L"foxz"

#define shTS					L"shTS"

#define HANGMUC					L"HM"
#define SYMSPECIAL				L".-/"
#define SYMSCHECKBOX			L"<ck!>"

#define CreateKeySettings		L"MFBGTlVYTmxkSFJwYm1kemZcEheF1lUeEGVRVTAJ51ZZa1ZjUjFoRVNsRVhOPURF"
#define TracuuVanban			L"TracuuVanban"

#define DlgSize					L"DlgSize"
#define DlgWidthColumn			L"DlgWidthColumn"

#define getSizeArray(arr)		sizeof(arr) / sizeof(arr[0]);

/*****************************************************************************************************/

static const int CONST_TIMER_CLISTEX	= 1;
static const int CONST_TIMER_TRAVANBAN	= 2;

static const CString scvtLaMa = L"IVXLCDM";
static const CString scvtNumber = L"0123456789";
static const CString scvtABC = L"ABCDEFGHIJKLMNOPQRSTUVWXYZ";

static const CString scvtLower = L"àáảãạăằắẳẵặâầấẩẫậòóỏõọôốồổỗộơờớởỡợèéẻẽẹêềếểễệùúủũụưừứửữựìíỉĩịỳýỷỹỵđ";
static const CString scvtUpper = L"ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÈÉẺẼẸÊỀẾỂỄỆÙÚỦŨỤƯỪỨỬỮỰÌÍỈĨỊỲÝỶỸỴĐ";

static const CString scvtUTF = L"àáảãạăằắẳẵặâầấẩẫậÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬòóỏõọôốồổỗộơờớởỡợÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢèéẻẽẹêềếểễệÈÉẺẼẸÊỀẾỂỄỆùúủũụưừứửữựÙÚỦŨỤƯỪỨỬỮỰìíỉĩịỳýỷỹỵđÌÍỈĨỊỲÝỶỸỴĐ";
static const CString scvtKOD = L"aaaaaaaaaaaaaaaaaAAAAAAAAAAAAAAAAAoooooooooooooooooOOOOOOOOOOOOOOOOOeeeeeeeeeeeEEEEEEEEEEEuuuuuuuuuuuUUUUUUUUUUUiiiiiyyyyydIIIIIYYYYYD";

static const string base64_chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

static const CString solama[] = {
	L"I",L"II",L"III",L"IV",L"V",L"VI",L"VII",L"VIII",L"IX",L"X",
	L"XI",L"XII",L"XIII",L"XIV",L"XV",L"XVI",L"XVII",L"XVIII",L"XIX",L"XX",
	L"XXI",L"XXII",L"XXIII",L"XXIV",L"XXV",L"XXVI",L"XXVII",L"XXVIII",L"XXIX",L"XXX",
	L"XXXI",L"XXXII",L"XXXIII",L"XXXIV",L"XXXV",L"XXXVI",L"XXXVII",L"XXXVIII",L"XXXIX",L"XL",
	L"XLI",L"XLII",L"XLIII",L"XLIV",L"XLV",L"XLVI",L"XLVII",L"XLVIII",L"XLIX",L"L",
	L"LI",L"LII",L"LIII",L"LIV",L"LV",L"LVI",L"LVII",L"LVIII",L"LIX",L"LX",
	L"LXI",L"LXII",L"LXIII",L"LXIV",L"LXV",L"LXVI",L"LXVII",L"LXVIII",L"LXIX",L"LXX",
	L"LXXI",L"LXXII",L"LXXIII",L"LXXIV",L"LXXV",L"LXXVI",L"LXXVII",L"LXXVIII",L"LXXIX",L"LXXX",
	L"LXXXI",L"LXXXII",L"LXXXIII",L"LXXXIV",L"LXXXV",L"LXXXVI",L"LXXXVII",L"LXXXVIII",L"LXXXIX",L"XC",
	L"XCI",L"XCII",L"XCIII",L"XCIV",L"XCV",L"XCVI",L"XCVII",L"XCVIII",L"XCIX",L"C" };

// 95 mã ASCII tương ứng từ 32 -> 126. Chú ý: Len=97 bao gồm 2 ký tự đặc biệt (\") & (\\) + 93 ký tự còn lại
static const CString szASCII = L" ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!\"#$%&‘()*+,–./:;<=>?@[\\]^_`{|}~";

// 95 kích thước Times New Roman chuẩn
static const double dWTimesNewRoman[] = { 0.2526,0.6876,0.6876,0.6876,0.7494,0.6252,0.5634,0.7494,0.7494,0.3144,0.4392,0.7494,0.6252,
	0.936,0.7494,0.7494,0.6252,0.7494,0.6876,0.5634,0.5634,0.7494,0.6876,0.9984,0.7494,0.6876,0.6252,0.4392,0.501,0.501,
	0.501,0.501,0.3768,0.501,0.501,0.3144,0.3144,0.501,0.3144,0.8118,0.501,0.5634,0.501,0.501,0.3768,0.4392,0.3144,0.501,
	0.4392,0.6876,0.501,0.501,0.4392,0.5634,0.5634,0.5634,0.5634,0.5634,0.5634,0.5634,0.5634,0.5634,0.5634,0.3144,0.3144,
	0.5634,0.5634,0.8742,0.8118,0.1902,0.3768,0.3768,0.5634,0.6252,0.2526,0.3768,0.2526,0.3144,0.3144,0.2526,0.6252,0.6252,
	0.6252,0.4392,0.936,0.3144,0.3144,0.3144,0.501,0.5634,0.3768,0.501,0.1902,0.501,0.5634 };