#pragma once
#include "StdAfx.h"

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

	void _msgupdate();

/*************** Excel ******************************************/

	// Hàm khởi tạo Excel
	// itype = 0 --> Chỉ gọi hàm | itype = 1 --> Gọi hàm có hiển thị hộp thoại
	void _xlCreateExcel(int itype = 0);

	// Hàm đóng Excel
	void _xlCloseExcel();

	// Hàm xác định thông tin sheet active
	void _xlGetInfoSheetActive();

	// Hàm kiểm tra và trả về phiên bản Excel đang sử dụng {9,10,..,16,..}
	// office2000 = 9 | xp=10 | 2003=11 | 2007=12 | 2010=14 | 2013=15 | 2016=16 |2019=16?
	CString _xlGetVersionExcel();

	// Tự động Enable Macro
	void _xlDeveloperMacroSettings();

	// Hàm kiểm tra Excel đang sử dụng là 32-bit hay 64-bit
	// Kết quả trả về giá trị = true nếu là 64-bit hoặc : = false nếu là 32-bit
	bool _xlIsOperatingSystem64();

	// Hàm cập nhật màn hình Excel
	// bl = false (Không cập nhật)
	// bl = true (Có cập nhật)
	void _xlPutScreenUpdating(bool bl = false);

	// Hàm chạy dòng statusbar Excel
	void _xlPutStatus(CString szStatus = L"Ready");

	// Hàm xác định Name bảng tính thông qua Codename
	// codename = codename bảng tính cần xác định
	// itype = 1 --> Thông báo nếu không tìm thấy sheet | itype = 0 --> bỏ qua không thông báo
	_bstr_t _xlGetNameSheet(_bstr_t codename, int itype = 0);

	// Hàm tìm kiếm giá trị 1 cột nào đó, nếu không tìm thấy trả về kết quả 0
	// pSheet = Bảng tính chứa name sử dụng
	// column = cột cần tìm giá trị & rowbd, rowkt = dòng bắt đầu & kết thúc tìm kiếm
	// match_content = true --> Tìm chính xác nội dung | false -> Chỉ cần xuất hiện cụm từ tìm kiếm
	// match_case = true --> Phân biệt chữ hoa chữ thường
	// ishow = 1 --> Trạng thại show cột nếu cột ẩn (ishow = 1 -> Hiện cột)
	// itype = 1 --> Thông báo nếu không tìm thấy giá trị | itype = 0 --> bỏ qua không thông báo
	int _xlFindValue(_WorksheetPtr pSheet, int column, int rowbd, int rowkt,
		_bstr_t bvalue, bool match_content, bool match_case, int ishow, int itype);

	// Hàm trả về giá trị tại ô chứa name được chỉ định trong sheet sử dụng
	// pSheet = phạm vi sheet sử dụng
	// name = tên name được chỉ định
	// itype = {0,1} | Nếu = 1 --> hiển thị thông báo nếu không tồn tại name
	// inumberformat = {0,1} | Nếu = 1 --> Nếu muốn lấy dữ liệu định dạng ngày tháng (ví dụ: 27/05/2017 --> 42882)
	CString _xlGetValue(_WorksheetPtr pSheet, CString name, int itype = 0, int inumberformat = 0);

	// Hàm đổ dữ liệu ra ô cell chứa Name (Kiểu chuỗi)
	// pSheet = Bảng tính chứa name sử dụng
	// name = Tên Name của ô cell cần đổ dữ liệu
	// szvalue = Giá trị chuỗi được sử dụng để đổ ra ô cell
	// itype = 1 --> Thông báo nếu không tìm thấy name | itype = 0 --> bỏ qua không thông báo
	void _xlPutValue(_WorksheetPtr pSheet, CString name, CString szvalue, int itype = 0);

	// Hàm đổ dữ liệu ra ô cell chứa Name (Kiểu số)
	// pSheet = Bảng tính chứa name sử dụng
	// name = Tên Name của ô cell cần đổ dữ liệu
	// ivalue = Giá trị số sẽ đổ ra ô cell
	// itype = 1 --> Thông báo nếu không tìm thấy name | itype = 0 --> bỏ qua không thông báo
	void _xlPutValue(_WorksheetPtr pSheet, CString name, int ivalue, int itype);

	// Hàm xác định vị trí dòng của Name được chọn
	// pSheet = Bảng tính chứa name sử dụng
	// name = Name cần xác định vị trí dòng
	// itype = 1 --> Thông báo mất name nếu không tìm thấy | itype = 0 --> bỏ qua không thông báo
	// itype = -1 sẽ return 0 --> Để lấy giá trị kiểm tra điều kiện ở ngoài
	int _xlGetRow(_WorksheetPtr pSheet, CString name, int itype);

	// Hàm xác định vị trí cột của Name được chọn
	// pSheet = Bảng tính chứa name sử dụng
	// name = Name cần xác định vị trí cột
	// itype = 1 --> Thông báo mất name nếu không tìm thấy | itype = 0 --> bỏ qua không thông báo
	// itype = -1 sẽ return 0 --> Để lấy giá trị kiểm tra điều kiện ở ngoài
	int _xlGetColumn(_WorksheetPtr pSheet, CString name, int itype);

	// Hàm lấy giá trị ô kiểu R1C1 (kể cả giá trị bị lỗi)
	// pRange = vùng tác động tại sheet lựa chọn
	// nRow, nColumn = cột, hàng sử dụng
	// szError = giá trị sẽ trả về nếu kết quả lấy bị lỗi (#REF, #VALUE,...)
	CString _xlGIOR(RangePtr pRange, int nRow, int nColumn, CString szError = L"");

	// Hàm chuyển đổi ô cell từ kiểu R1C1 sang A1
	// pRange = vùng tác động tại sheet lựa chọn
	// nRow, nColumn = cột, hàng sử dụng
	// itype = kiểu dữ liệu trả về (itype=0 --> A1 | itype=1 $A1 | itype=2 A$1 | itype=3 $A$1)
	CString _xlGAR(RangePtr pRange, int nRow, int nColumn, int itype);

	// Hàm chuyển đổi định dạng từ R1C1 sang A1 (Ví dụ: Cột 20 --> T)
	// iColumn = cột kiểu R1C1 --> chuyển sang A1
	CString _xlConvertRCtoA1(int iColumn);

	// Messagebox thông báo mất name
	void _xlMsgNotName(_WorksheetPtr pSheet, CString name);

	// Hàm tự động khôi phục name nếu như không tìm thấy
	void _xlCreateNewName(CString szName, CString &szKpname, int &iKprow, int &iKpcol);

	// Set hyperlink
	// pSheet = Sheet sử dụng để set
	// PRS = Ô dữ liệu sử dụng để set
	// pathFile = Đường dẫn gắn vào địa chỉ liên kết
	// szName = Tên hiện thị
	// clrBkg, clrTxt = Màu nền và chữ hiển thị
	// szFontName = Font chữ | iSize = Kích thước chữ | bCenter = Căn giữa (hoz)
	void _xlSetHyperlink(_WorksheetPtr pSheet, RangePtr PRS, 
		CString pathFile, CString szName = L"Mở file",
		COLORREF clrBkg = xlNone, COLORREF clrTxt = rgbBlue, 
		CString szFontName = FontArial, int iSize = 12, 
		bool bCenter = true, bool bFormula = true);

	// Hàm lấy tất cả các sheet
	// Giá trị trả về là tổng số sheet khác VeryHidden (2) và lưu vào vector <wpSheet>
	// nHidden = Số lượng sheet ẩn
	int _xlGetAllSheet(vector<_WorksheetPtr> &wpSheet, int &nHidden);

	// Hàm xác định vị trí tương ứng của pSheet
	int _xlGetIndexSheet(CString szCodename);

	// Hàm tạo mới sheet, trả về _WorksheetPtr nếu tạo thành công (true)
	// nIndex = Vị trí sheet chỉ định đứng trước sheet được tạo, nếu = null --> sheet được tạo ở vị trí cuối cùng
	// szCodename, szName = Code và name của sheet được tạo
	// clrTab = Màu tab được tạo
	bool _xlCreateNewSheet(_WorksheetPtr &wPsh, int nIndex = 0, CString szCodename = L"", CString szName = L"", COLORREF clrTab = xlAutomatic);

	// Hàm lấy giá trị comment tại pRange = Ô dữ liệu được sử dụng
	CString _xlGetcomment(RangePtr pRange);

	// Hàm đổ comment <szComment> tại pRange = Ô dữ liệu được sử dụng
	void _xlPutcomment(RangePtr pRange, CString szComment);

	// Hàm định dạng khung vùng dữ liệu
	// iStyle = Kiểu dữ liệu định dạng riêng cho xlInsideHorizontal (iStyle = 1 xlThin | 2 = xlHairline | 3 = xlDot)
	// blTop, blBottom = true/false = cho phép định dạng khung top và bottom
	void _xlFormatBorders(RangePtr pRange, int iStyle = 3, bool blTop = true, bool blBottom = true);

	// Hàm lấy liên kết hyperlink tại 1 ô duy nhất
	// iType = 1 --> replace '/' -> '\' | iType = 2 --> replace '\' -> '/'
	CString _xlGetHyperLink(RangePtr pRgSelect, int iType = 1);

	// Hàm xác định số lượng hyperlink trong vùng được chọn và lưu vào vector 'vecHyper'
	// iType = 1 --> replace '/' -> '\' | iType = 2 --> replace '\' -> '/'
	// iStatus = 1 --> Cập nhật thanh statusbar
	int _xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper, int iType = 1, int iStatus = 0);

	// Hàm xác định vùng lựa chọn sheet active
	CString _xlCheckSelection();

	// Hàm xác định chi tiết dữ liệu được chọn
	// Hàm trả về vector <vecSelect> - biến toàn cục
	// Bao gồm các trường: int= {row, column}, Cstring= {cell, value, formula};
	// itypeCell : 0= Lấy tất cả các ô dữ liệu | 1= Chỉ lấy các ô có dữ liệu | 2= Chỉ lấy các ô NULL
	// itypeArea : 0= Lấy tất cả vùng dữ liệu | 1= Chỉ lấy vùng dữ liệu được chọn đầu tiên
	int _xlGetAllCellSelection(int itypeCell = 0, int itypeArea = 0);

	// Hàm lưu các giá trị ô được chọn
	// iRow,iColumn = Cột, hàng vị trí ô cần lưu
	// szChuoi = Chuỗi được lưu
	// szFormula = Công thức sẽ lưu (nếu có)
	// itypeCell= 0 --> Lấy cả các ô không có dữ liệu | itypeCell= 1 Chỉ lấy các ô có dữ liệu
	void _xlSaveTotalSelection(int iRow, int iColumn, CString szChuoi, CString szFormula, int itypeCell);

	// Hàm xác định cột dòng được chọn
	// szPkey = giá trị thứ (i) của biến vecKey[i= 0..n]
	// RC có giá trị = R (row) hoặc C (column)
	// iRCEnd = vị trí cuối cùng chứa dữ liệu của dòng hoặc cột
	// itypeCell= 0 --> Lấy cả các ô không có dữ liệu | itypeCell= 1 Chỉ lấy các ô có dữ liệu
	void _xlSaveR1C1(CString szPkey, CString RC, int iRCEnd, int itypeCell);

	// Hàm gọi hàm từ DLL khác
	// szDllName = Tên DLL sử dụng
	// szFunction = Tên hàm trong DLL được gọi ra sử dụng
	// blQuit = true --> Đóng DLL sau khi kết thúc sử dụng hàm
	void _xlCallFunctionDLL(CString szDllName, CString szFunction, bool blQuit = true);	

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
	int _FileExistsUnicode(CString pathFile, int itype = 0);

	// Hàm kiểm tra có phải thư mục không?
	bool _FolderExistsUnicode(CString dirName_in);

	// Hàm lựa chọn file trong thư mục
	// szDirFileSelect = Đường dẫn chứa tệp tin được lựa chọn 
	void _ShellExecuteSelectFile(CString szDirFileSelect);

	// Hàm tạo mới 1 thư mục được chỉ định
	// Có thể bổ trợ cho hàm _CreateDirectories
	BOOL _CreateNewFolder(CString directoryPath);

	// Hàm tạo nhiều thư mục con từ đường dẫn được chỉ định
	bool _CreateDirectories(CString pathName);

	// Hàm không chính xác khi thư mục được tạo có sử dụng tính năng chia sẻ (Dropbox, Google,...)
	// Hàm xác định tất cả các tệp tin trong thư mục được chỉ định
	// Kết quả trả về là tổng số files tìm thấy và lưu trữ thông tin vào '_finder' (_finder.GetFilePath(i).GetFileName()...)
	// pathFolder = Đường dẫn thư mục tìm kiếm
	// typeOfFile = Kiểu file cần tìm kiếm (xls,mdb,...)
	int _FindAllFile(CFileFinder &_finder, CString pathFolder, CString typeOfFile = L"*");

	// Hàm xác định tất cả các tệp tin trong thư mục được chỉ định
	// Kết quả trả về là tổng số files tìm thấy và lưu trữ thông tin vào 'vecPathFiles'
	// szPathCha = Đường dẫn thư mục tìm kiếm
	// vecPathFiles = Vector sử dụng để lưu trữ thông tin
	int _FindAllFileEx(CString szPathCha, vector<CString> &vecPathFiles);

	// Hàm xác định tất cả các thư mục trong thư mục được chỉ định
	// Kết quả trả về là tổng số thư mục tìm thấy và lưu trữ thông tin vào 'vecPathFolder'
	// szPathCha = Đường dẫn gốc sử dụng để tìm kiếm thư mục
	// vecPathFolder = Vector sử dụng để lưu trữ thông tin
	int _FindAllFolder(CString szPathCha, vector<CString> &vecPathFolder);

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

	// Hàm sắp xếp nhanh mảng số
	// vec = Mảng số cần sắp xếp gồm N giá trị
	// ll = vị trí đầu tiên trong mảng (bắt đầu = 0)
	// rr = vị trí cuối cùng (N - 1)
	void _QuickSortArr(vector<int> &vec, int L, int R);

	// Hàm xác định đường dẫn Desktop
	CString _GetDesktopDir();

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

	// Hàm ghi dữ liệu lên file unicode
	// vecTextLine = các giá trị lần lượt được ghi
	// spathFile = Đường dẫn file sử dụng (Nếu không có sẽ tự tạo file)
	// itype = 1 --> Nếu không tồn tại file dừng lại luôn
	void _WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype = 0);

	// Hàm đọc dữ liệu unicode từ file text, kết quả trả về là tổng số dòng tìm thấy và lưu danh sách vào vector
	// spathFile = Đường dẫn tệp tin txt hoặc ini
	// vecTextLine = vector lưu trữ dữ liệu các dòng tìm được
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);

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

	// Hàm chuyển đổi sang chữ HOA có chứa ký tự Unicode
	CString _UpperUnicode(CString szNoidung);

	// Hàm chuyển đổi sang chữ thường có chứa ký tự Unicode
	CString _LowerUnicode(CString szNoidung);

	// Chỉ nên áp dụng trong trường hợp chuỗi có ký tự tiếng việt có dấu
	// Hàm Replace không phân biệt chữ HOA - thường
	CString _ReplaceMatchCase(CString szNoidung, CString szOld, CString szNew, bool blCase = true, bool blKDau = false);

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

	// Hàm kiểm tra có đăng ký được COM vào máy tính không?
	bool Function::_CheckRegisterCOM();

	// Hàm lấy thông tin từ file document có đường dẫn 'szpath'
	// nIndex = 0 --> Lấy tất cả thông tin
	// nIndex = 1: Title | 2: Subject | 3: Tags (Keywords) | 4: Categories
	// nIndex = 5: Comments | 6: Authors | 7: Company
	CString _GetProperties(CString szpath, int nIndex = 1);

	// Hàm ghi thông tin vào file document có đường dẫn 'szpath'
	// nIndex = 1: Title | 2: Subject | 3: Tags (Keywords) | 4: Categories
	// nIndex = 5: Comments | 6: Authors | 7: Company
	bool _SetProperties(CString szpath, int nIndex, CString strVal = L"");

	// Hàm ghi toàn bộ thông tin vào file document có đường dẫn 'szpath'
	bool _SetAllProperties(
		CString szpath,	CString szTitle = L"", CString szSubject = L"", CString szTags = L"", 
		CString szCategories = L"",	CString szComments = L"", CString szAuthor = L"", CString szCompany = L"");

	// Hàm xác định windows 32 hay 64. Kết quả trả về True nếu là windows 64bit 
	bool _Is64BitWindows();

/*** Phần hàm viết riêng **********************************************************************/

	// Hàm lấy thông tin sheet Category gốc
	void _GetThongtinSheetCategory(int &icotSTT, int &icotLV, int &icotMH, int &icotNoidung,
		int &icotFilegoc, int &icotFileGXD, int &icotThLuan, int &icotEnd, int &irowTieude, int &irowStart, int &irowEnd);

	// Hàm lấy thông tin sheet FManager gốc
	void _GetThongtinSheetFManager(int &icotSTT, int &icotLV, int &icotFolder, int &icotTenFile, int &icotType,
		int &icotNoidung, int &icotFilegoc, int &icotEnd, int &irowTieude, int &irowStart, int &irowEnd);

	// Hàm định nghĩa lại sheet active = szSheet
	void _DefaultSheetActive(CString szSheet);

	// Hàm xác định đã đăng ký COM cho file 'Gproperties' chưa
	void _RegistryCOMGpro();

	// Hàm hỗ trợ mở file (Chỉ áp dụng cho các sheet biểu mẫu hoặc quản lý files)
	void _xlOpenFileEx(_WorksheetPtr pSheet, int nRowActive, int nColumnActive);

	// Hàm hỗ trợ mở thư mục (Chỉ áp dụng cho các sheet biểu mẫu hoặc quản lý files)
	void _xlOpenFolderEx(_WorksheetPtr pSheet, int nRowActive, int nColumnActive);
};
