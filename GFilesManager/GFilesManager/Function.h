#pragma once
#include "pch.h"

typedef int(__stdcall *MSGBOXAAPI)(IN HWND hWnd, IN LPCSTR lpText, IN LPCSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

typedef int(__stdcall *MSGBOXWAPI)(IN HWND hWnd, IN LPCWSTR lpText, IN LPCWSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

#ifdef UNICODE
#define MessageBoxTimeout  MessageBoxTimeoutW
#else
#define MessageBoxTimeout  MessageBoxTimeoutA
#endif

int MessageBoxTimeoutA(IN HWND hWnd, IN LPCSTR lpText, IN LPCSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

int MessageBoxTimeoutW(IN HWND hWnd, IN LPCWSTR lpText, IN LPCWSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds);

typedef struct file_info_struct
{
	DWORD    dwFileAttributes;
	FILETIME ftCreationTime;
	FILETIME ftLastAccessTime;
	FILETIME ftLastWriteTime;
	DWORD    nFileSizeHigh;
	DWORD    nFileSizeLow;
} FILE_INFO;

typedef BOOL(WINAPI *GET_FILE_ATTRIBUTES_EX)(LPCWSTR lpFileName, int fInfoLevelId, LPVOID lpFileInformation);

class Function
{

public:
	Function();
	~Function();

/*************** Message box ************************************/
	// itype: 0 = MB_ICONINFORMATION | 1 = MB_ICONWARNING | 2 = MB_ICONERROR
	void _msgbox(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _s(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _d(int num = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _n(CString message = L"Num", int num = 0, int itype = 0, int istr = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	int _y(CString message = L"", int istyle = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);

/*************** Functions ****************************************/
	
	// Checkversion Software
	int _CheckVersionSoftware();

	// Hàm chuyển từ kiểu Cstring --> char*
	char* _ConvertCstringToChars(CString szvalue);

	// Hàm chuyển từ kiểu char* --> Cstring
	CString _ConvertCharsToCstring(char* chr);

	// Hàm kiểm tra tệp tin có tồn tại hay không (return = 1 --> tồn tại)
	// pathFile = đường dẫn tệp tin cần kiểm tra
	// itype = 1 thông báo nếu có lỗi xảy ra | itype = 0 bỏ qua thông báo
	int _FileExists(CString pathFile, int itype = 1);

	// Hàm kiểm tra tệp tin hoặc thư mục (bao gồm cả Unicode) có tồn tại hay không (return = 1 --> tồn tại)
	// pathFile = đường dẫn tệp tin cần kiểm tra
	// itype = 1 thông báo nếu có lỗi xảy ra | itype = 0 bỏ qua thông báo
	int _FileExistsUnicode(CString pathFile, int itype);

	// Hàm kiểm tra có phải thư mục không?
	bool _FolderExistsUnicode(CString dirName_in);

	// Hàm tạo mới 1 thư mục được chỉ định
	// Có thể bổ trợ cho hàm _CreateDirectories
	BOOL _CreateNewFolder(CString directoryPath);

	// Hàm tạo nhiều thư mục con từ đường dẫn được chỉ định
	bool _CreateDirectories(CString pathName);

	// Hàm xác định tất cả các tệp tin trong thư mục được chỉ định
	// Kết quả trả về là tổng số files tìm thấy và lưu trữ thông tin vào '_finder' (_finder.GetFilePath(i).GetFileName()...)
	// pathFolder = Đường dẫn thư mục tìm kiếm
	// typeOfFile = Kiểu file cần tìm kiếm (xls,mdb,...)
	int _FindAllFile(CFileFinder &_finder, CString pathFolder, CString typeOfFile = L"*");

	// Hàm lấy đường dẫn của tệp tin 'fileName' được sử dụng
	CString _GetPathFolder(CString fileName);

	// Hàm kiểm tra ứng dụng có chạy hay không (Ví dụ: GFilesManager.exe)
	bool _IsProcessRunning(const wchar_t *processName);

	// Hàm tách từ khóa chính thành các từ khóa con. Kết quả trả về là số lượng và chi tiết từ khóa con tìm được
	// vecPkey = vector lưu trữ danh sách từ khóa con
	// sKeyValue = Từ khóa chính
	// symbol1, symbol2 = Ký tự sử dụng để phân cách các từ khóa (symbol2 có thể bằng chính symbol1)
	// itypeFilter = 1 --> Lọc các từ khóa bị trùng lặp | itypeFilter = 0 --> Lấy tất cả từ khóa
	// itypeTrim = 1 --> Xóa khoảng trắng 2 đầu mỗi từ khóa | itypeFilter = 0 --> Giữ nguyên từ khóa không làm gì cả
	int _TackTukhoa(vector<CString> &vecPkey, CString sKeyValue, CString symbol1 = L";", CString symbol2 = L"|", int itypeFilter = 0, int itypeTrim = 1);

	// Hàm lấy số ngẫu nhiên theo thời gian
	int _RandomTime();

	// Hàm lấy thời gian thực theo định dạng dd/MM/yyyy
	// ifulltime = 1 --> Trả vể đầy đủ cả ngày giờ dd/MM/yyyy HH:mm:ss
	CString _GetTimeNow(int ifulltime = 0);

	// Hàm so sánh 2 ngày dateBefore (trước) & dateAfter (sau)
	// Giá trị trả về: true = dateBefore <= dateAfter | hoặc false = dateBefore > dateAfter
	bool _CompareDate(CString dateBefore, CString dateAfter);

	// Hàm cộng thêm [num] ngày tính từ ngày 'szdate' được chỉ định
	CString	_DayNextPrev(CString szdate, int num = 1);

	// Hàm trả về đuôi file dữ liệu được chỉ định
	// pathFile = Đường dẫn dữ liệu
	CString _GetTypeOfFile(CString pathFile);

	// Hàm trích xuất tên file hoặc tên đường dẫn
	// pathFile = Đường dẫn dữ liệu sử dụng
	// pos = Giá trị được trả về là vị trí phân tách
	// ipath= -1 --> Chỉ trích xuất tên file  
	// ipath=  0 --> Trích xuất tên file + đuôi 
	// ipath=  1 --> Trích xuất tên đường dẫn chứa file
	CString _GetNameOfPath(CString pathFile, int &pos, int ipath = 0);

	// Hàm xác định đường dẫn Desktop
	CString _GetDesktopDir();

	// Hàm ghi dữ liệu lên file unicode
	// vecTextLine = các giá trị lần lượt được ghi
	// spathFile = Đường dẫn file sử dụng (Nếu không có sẽ tự tạo file)
	// itype = 1 --> Nếu không có dừng lại luôn
	void _WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype = 0);

	// Hàm đọc dữ liệu unicode từ file text, kết quả trả về là tổng số dòng tìm thấy và lưu danh sách vào vector
	// spathFile = Đường dẫn tệp tin txt hoặc ini
	// vecTextLine = vector lưu trữ dữ liệu các dòng tìm được
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);

	// Hàm xác định tên ứng dụng đang chạy trên Process
	CString _GetDirModules(DWORD processID, bool blFilter = false);

	// Hàm ghi lại dữ liệu cần thiết kiểu chuỗi(Lỗi, ghi chú,...)
	// blClear = true --> Xóa toàn bộ dữ liệu trước khi ghi kết quả
	// blOpenFile = true --> Mở file để xem kết quả
	// szFileName = Tên file sẽ sử dụng để đổ dữ liệu (Nếu chưa có sẽ tự tạo mới)
	void _LogFileWrite(CString szLog = L"", bool blClear = false, bool blOpenFile = false, CString szFileName = FileLog);

	// Hàm ghi lại dữ liệu cần thiết kiểu số (Lỗi, ghi chú,...)
	// blClear = true --> Xóa toàn bộ dữ liệu trước khi ghi kết quả
	// blOpenFile = true --> Mở file để xem kết quả
	void _LogFileWrite(int iLog = 0, bool blClear = false, bool blOpenFile = false, CString szFileName = FileLog);

	// Hàm bắt đầu ghi log
	void _LogFileStart(CString szStart = L"<!>");

	// Hàm kết thúc ghi log
	void _LogFileEnd();

	// Hàm convert không dấu
	// szNoidung = Nội dung cần convert
	CString _ConvertKhongDau(CString szNoidung);

	// Hàm đổi giá trị từ double = ivalue (Bytes) sang KB,MB,GB,TB,...
	CString _ConvertBytes(double ivalue);

	// Hàm làm tròn số thập phân
	// doValue = Giá trị cần làm tròn
	// nPrecision = Làm tròn sau bao nhiêu số thập phân
	double _RoundDouble(double doValue, int nPrecision);

	// Hàm xác định thông tin file (bao gồm kích thước và ngày)
	// szDir = Đường dẫn thư mục cần tìm kiếm tệp tin
	// Giá trị trả về lưu trữ = szKichthuoc thước & szNgay
	void _GetFileAttributesEx(CString szDir, CString &szKichthuoc, CString &szNgay);

	// Hàm format định dạng tiền VNĐ
	// szTien = Mệnh giá (giá trị là số)
	// szSteparator = Ký tự ngăn cách giữa hàng nghìn, trăm, ... 
	// szDecimal = Ký tự ngăn cách với hàng thập phân
	CString _FormatTienVND(CString szTien, CString szSteparator, CString szDecimal);

	// Hàm convert tên thư mục/ sheet
	CString _ConvertRename(CString sname);

	// Convert chuỗi sang màu COLORREF
	// itype = 0 --> RGBWHITE | itype = 1 --> RGBBLACK;
	COLORREF _SetColorRGB(CString szColor, int itype = 0);

	// Convert màu COLORREF sang chuỗi
	CString _GetColorRGB(COLORREF clr);

	// Hàm lựa chọn file trong thư mục
	// szDirFileSelect = Đường dẫn chứa tệp tin được lựa chọn 
	void _ShellExecuteSelectFile(CString szDirFileSelect);

	// Hàm gọi tới 1 dll khác
	void _CallFunctionDLL(CString szDllName, CString szFunction);

	// Hàm xác định windows 32 hay 64. Kết quả trả về True nếu là windows 64bit 
	bool _Is64BitWindows();	
};
