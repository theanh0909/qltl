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

class CFunction
{

public:
	CFunction();
	~CFunction();

/*************** Message box ************************************/

	// itype: 0 = MB_ICONINFORMATION | 1 = MB_ICONWARNING | 2 = MB_ICONERROR
	void _msgbox(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _s(CString message = L"", int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _d(int num = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	void _n(CString message = L"Value", int num = 0, int itype = 0, int istr = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);
	int _y(CString message = L"", int istyle = 0, int itype = 0, int iexcel = 0, int itimeout = MINISEC_TIME_OUT);

	void _msgupdate();

/***********************************************************************************************/

	// Hàm khởi tạo Excel
	// blOpenFrm = true --> Trong hàm được sử dụng có mở Form
	bool _xlCreateExcel(bool blOpenFrm = false);

	// Hàm xác định phiên bản Excel đang sử dụng
	CString _xlGetVersionExcel();

	// Hàm Enable devloper macro settings
	void _xlDeveloperSettings();

	// Hàm kiểm tra Excel đang sử dụng là 32-bit hay 64-bit
	// Kết quả trả về giá trị = true nếu là 64-bit hoặc : = false nếu là 32-bit
	bool _xlIsOperatingSystem64();
	
	// Hàm chạy dòng statusbar Excel
	void _xlPutStatus(CString szStatus = L"Ready");

	// Hàm xác định Name của sheet dựa vào codename 
	// blError = True --> Hiển thị thông báo lỗi nếu không tìm thấy sheet
	_bstr_t _xlGetNameSheet(_bstr_t codename, bool blError = false);

	// Hàm xác định vị trí dòng của Name được chọn
	// pSheet = Bảng tính chứa name sử dụng
	// name = Name cần xác định vị trí dòng
	// itype = 1 --> Thông báo mất name nếu không tìm thấy | itype = 0 --> bỏ qua không thông báo
	// itype = -1 sẽ return 0 --> Để lấy giá trị kiểm tra điều kiện ở ngoài
	int _xlGetRow(_WorksheetPtr pSheet, CString name, int itype = 1);

	// Hàm xác định vị trí cột của Name được chọn
	// pSheet = Bảng tính chứa name sử dụng
	// name = Name cần xác định vị trí cột
	// itype = 1 --> Thông báo mất name nếu không tìm thấy | itype = 0 --> bỏ qua không thông báo
	// itype = -1 sẽ return 0 --> Để lấy giá trị kiểm tra điều kiện ở ngoài
	int _xlGetColumn(_WorksheetPtr pSheet, CString name, int itype);

	// Hàm xác định cột cuối cùng của bảng dữ liệu
	// rowGet = Vị trí dòng được sử dụng để xác định
	// icolStart = Vị trí cột bắt đầu xác định
	// icolEnd = Vị trí kết thúc xác định, nếu icolEnd = 0 --> Lấy icolEnd = vị trí cuối cùng chứa dữ liệu
	// itype = 0 --> Lấy vị trí cuối cùng chứa xlThin (|) tìm được
	// itype = 1 --> Lấy vị trí đầu tiên chứa xlThin (|) tìm được
	int _xlGetTableColumnLast(_WorksheetPtr pSheet, int rowGet, int icolStart = 1, int icolEnd = 0, int itype = 0);

	// Hàm xác định dòng cuối cùng của bảng dữ liệu
	// colGet = Vị trí dòng được sử dụng để xác định
	// irowStart = Vị trí cột bắt đầu xác định
	// irowEnd = Vị trí kết thúc xác định, nếu icolEnd = 0 --> Lấy icolEnd = vị trí cuối cùng chứa dữ liệu
	// itype = 0 --> Lấy vị trí cuối cùng chứa xlThin (|) tìm được
	// itype = 1 --> Lấy vị trí đầu tiên chứa xlThin (|) tìm được
	int _xlGetTableRowLast(_WorksheetPtr pSheet, int colGet, int irowStart = 1, int irowEnd = 0, int itype = 0);

	// Hàm trả về tổng độ rộng các cột cộng lại
	// pRange = Vùng dữ liệu cần tính
	// iColumnStart,iColumnEnd = Cột bắt đầu và kết thúc tính
	double _xlTotalWidthColumn(RangePtr pRange, int iColumnStart, int iColumnEnd);

	// Hàm trả về độ dài chuỗi văn bản trong Excel
	// szNoidung = Nội dung chuỗi cần kiểm tra
	// szFont = Font chữ sử dụng tại vị trí đó (Hiện code chỉ hỗ trợ font Times New Roman)
	// iSize = Kích thước font được sử dụng
	double _xlGetWidthStringExcel(CString szNoidung, CString szFont = FontTimes, double iSize = 12);

	// Hàm tự động Merge và Autofit vùng dữ liệu được chọn (Update 24.12.2018)
	// pSheet = Bảng tính chứa name sử dụng
	// iRowStart = Dòng sử dụng dữ liệu cần gộp
	// iColumnStart, iColumnEnd = Cột bắt đầu và kết thúc gộp
	// iTotalW = Tổng độ rộng các cột cộng lại (iColumnStart --> iColumnEnd)
	// itype = 1 --> Cho phép so sánh chiều cao sẽ áp dụng với chiều cao cũ, được sử dụng khi phải autofit nhiều cột
	void _xlAutoMergeCellNew(_WorksheetPtr pSheet, int iRowStart, int iColumnStart, int iColumnEnd, double iTotalW, int itype = 0);

	// Hàm set hyperlink
	// pSheet = Sheet sử dụng để set
	// pCell = Ô dữ liệu sử dụng để set
	// pathFile = Đường dẫn gắn vào địa chỉ liên kết
	// szName = Tên hiện thị
	// clrBkg, clrTxt = Màu nền và chữ hiển thị
	// szFontName = Font chữ | iSize = Kích thước chữ | bCenter = Căn giữa (hoz)
	void _xlSetHyperlink(_WorksheetPtr pSheet, RangePtr pCell,
		CString pathFile, CString szName = L"Mở file",
		COLORREF clrBkg = xlNone, COLORREF clrTxt = rgbBlue,
		CString szFontName = FontArial, int iSize = 12,
		bool bCenter = true, bool bFormula = true);

	// Hàm lấy liên kết hyperlink tại 1 ô duy nhất
	// iType = 1 --> replace '/' -> '\' | iType = 2 --> replace '\' -> '/'
	CString _xlGetHyperLink(RangePtr pRgSelect, int iType = 1);

	// Hàm xác định số lượng hyperlink trong vùng được chọn và lưu vào vector 'vecHyper'
	// iType = 1 --> replace '/' -> '\' | iType = 2 --> replace '\' -> '/'
	// iStatus = 1 --> Cập nhật thanh statusbar
	int _xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper, int iType = 1, int iStatus = 0);

	// Hàm tự động khôi phục name nếu như không tìm thấy
	void _xlCreateNewName(CString szName, CString &szKpname, int &iKprow, int &iKpcol);

	// Hàm thông báo mất name
	void _xlMsgNotName(_WorksheetPtr pSheet, CString name);

	// Hàm lấy giá trị ô kiểu R1C1 (kể cả giá trị bị lỗi)
	// pRange = vùng tác động tại sheet lựa chọn
	// nRow, nColumn = cột, hàng sử dụng
	// szError = giá trị sẽ trả về nếu kết quả lấy bị lỗi (#REF, #VALUE,...)
	CString _xlGIOR(RangePtr pRange, int nRow, int nColumn, CString szError = L"");

	// Hàm tìm kiếm Comment tại 1 cột nào đó
	// pSheet = Bảng tính chứa name sử dụng
	// column = Cột chứa comment cần tìm kiếm
	// rowStart, rowEnd = Dòng bắt đầu và kết thúc tìm kiếm
	// comment = comment cần tìm kiếm
	// itype = 1 --> Thông báo nếu không tìm thấy comment | itype = 0 --> bỏ qua không thông báo
	// itype = -1 sẽ return 0 --> Để lấy giá trị kiểm tra điều kiện ở ngoài 
	int _xlFindComment(_WorksheetPtr pSheet, int column = 1, 
		int rowStart = 0, int rowEnd = 0, _bstr_t comment = (_bstr_t)L"END", int itype = 0);

	// Hàm chuyển đổi ô cell từ kiểu R1C1 sang A1
	// pRng = vùng tác động tại sheet lựa chọn
	// nRow, nColumn = cột và hàng tại ô được chọn
	// itype = kiểu dữ liệu trả về (itype=0 --> A1 | itype=1 $A1 | itype=2 A$1 | itype=3 $A$1)
	CString _xlGAR(RangePtr pRng, int nRow, int nColumn, int itype = 0);

	// Hàm chuyển đổi định dạng từ A1 sang R1C1 (Ví dụ: T12 --> R20C12)
	// szValueA1 = giá trị kiểu A1 --> chuyển sang R1C1. 
	// Nếu giá trị chỉ là A,B,.. --> Mặc định được đổi thành A1,B1,...
	// nRow, nColumn = Kết quả trả về sẽ nhận được
	bool _xlConvertA1toRC(CString szValueA1, int &nRow, int &nColumn);

	// Hàm chuyển đổi định dạng từ R1C1 sang A1 (Ví dụ: Cột 20 --> T)
	// iColumn = cột kiểu R1C1 --> chuyển sang A1
	CString _xlConvertRCtoA1(int iColumn);

	// Hàm xóa định dạng vùng dữ liệu
	void _xlFormatClearFormat(RangePtr pRange);

	// Hàm định dạng khung vùng dữ liệu
	// iStyle = Kiểu dữ liệu định dạng riêng cho xlInsideHorizontal (iStyle = 1 xlThin | 2 = xlHairline | 3 = xlDot)
	// blTop, blBottom = true/false = cho phép định dạng khung top và bottom
	void _xlFormatBorders(RangePtr pRange, int iStyle = 3, bool blTop = true, bool blBottom = true);

	// Hàm cập nhật màn hình Excel
	// bl = false (Không cập nhật) hoặc bl = true (Có cập nhật)
	void _xlPutScreenUpdating(bool bl = false);

	// Hàm get comment 
	CString _xlGetcomment(RangePtr pRange);

	// Hàm add comment
	// szComment = Nội dung comment
	void _xlPutcomment(RangePtr pRange, CString szComment = L"END");

	// Hàm xác định giá trị ô cell theo Name
	// pSheet = Bảng tính chứa name sử dụng
	// name = Name tại vị trí ô cell
	// itype = 1 --> Thông báo mất name nếu không tìm thấy | itype = 0 --> bỏ qua không thông báo
	// inumberformat = 1 --> Kiểu lấy dữ liệu định dạng đặc biệt (ví dụ: ngày tháng: 27/05/2017 --> 42882)
	CString _xlGetValue(_WorksheetPtr pSheet, CString name, int itype = 1, int inumberformat = 0);

	// Hàm đổ dữ liệu ra ô cell chứa Name
	// pSheet = Bảng tính chứa name sử dụng
	// name = Tên Name của ô cell cần đổ dữ liệu
	// szValue = Giá trị sẽ đổ ra ô cell
	// itype = 1 --> Thông báo nếu không tìm thấy name | itype = 0 --> bỏ qua không thông báo
	void _xlPutValue(_WorksheetPtr pSheet, CString name, CString szValue, int itype = 1);

	// Hàm đổ dữ liệu ra ô cell chứa Name
	// pSheet = Bảng tính chứa name sử dụng
	// name = Tên Name của ô cell cần đổ dữ liệu
	// iValue = Giá trị sẽ đổ ra ô cell
	// itype = 1 --> Thông báo nếu không tìm thấy name | itype = 0 --> bỏ qua không thông báo
	void _xlPutValue(_WorksheetPtr pSheet, CString name, int iValue, int itype = 1);

	// Hàm add InputMessage trong Validation
	// szNoidung = Nội dung cần hiển thị hướng dẫn
	void _xlPutInputMessage(RangePtr pRange, CString szNoidung, CString szHuongdan = L"Hướng dẫn");

	// Hàm xác định và lấy thông tin sheet active
	void _xlGetInfoSheetActive();

	// Hàm tạo mới sheet, trả về _WorksheetPtr nếu tạo thành công (true)
	// nIndex = Vị trí sheet chỉ định đứng trước sheet được tạo, nếu = null --> sheet được tạo ở vị trí cuối cùng
	// szCodename, szName = Code và name của sheet được tạo
	// clrTab = Màu tab được tạo
	bool _xlCreateNewSheet(_WorksheetPtr &wPsh, int nIndex = 0, 
		CString szCodename = L"", CString szName = L"", COLORREF clrTab = xlAutomatic);
	
	// Hàm lấy tất cả các sheet
	// Giá trị trả về là tổng số sheet khác VeryHidden (2) và lưu vào vector <wpSh>
	// nVeryHidden = Số lượng sheet ẩn hoàn toàn
	int _xlGetAllSheet(vector<_WorksheetPtr> &wpSh, int &nVeryHidden);

	// Hàm xác định vị trí tương ứng của pSheet
	// Các sheet khác VeryHidden (2) được lưu vào vector <wpSh>
	int _xlGetIndexSheet(CString szCodename, vector<_WorksheetPtr> &vecPsh);

	// Hàm tạo Name & Codename sheet mới
	// vecPsh = Danh sách toàn bộ sheet khác VeryHidden (2)
	bool _xlCreateNewNameSheet(CString &szName, CString &szCodename, vector<_WorksheetPtr> vecPsh);

	// Hàm xác định đậm tên font chữ
	CString _xlGetCellFontName(RangePtr pCell);

	// Hàm xác định đậm ô cell
	int _xlGetCellSize(RangePtr pCell);

	// Hàm xác định đậm ô cell
	bool _xlGetCellBold(RangePtr pCell);

	// Hàm xác định kiểu nghiêng ô cell
	bool _xlGetCellItalic(RangePtr pCell);

	// Hàm xác định màu chữ ô cell
	COLORREF _xlGetCellTxtColor(RangePtr pCell);

	// Hàm xóa toàn bộ shape tại sheet được chỉ định
	// sName = Tên shape sẽ được xóa, nếu sName = NULL --> Xóa toàn bộ
	// blEntire = true --> Tìm kiếm chính xác tên, ngược lại chỉ lấy LEFT 
	// Kết quả trả về là tổng số lượng shape được xóa
	int _xlDeleteShape(_WorksheetPtr pSheet, CString sName, bool blEntire = false);

	// Hàm xác định kích thước ảnh
	// szPathIMG = Đường dẫn hình ảnh sử dụng
	// x,y - Kích thước ảnh thu về trả lại giá trị &qImageW + &qImageH
	bool _GetImageSize(CString szPathIMG, int *x, int *y);

	// Hàm tự động tạo tên shape trên 1 sheet chỉ định
	// szPrefix = Tiền tố sử dụng (nếu có)
	CString _xlCreateNewNameShape(_WorksheetPtr pSheet, CString szPrefix = L"GXD");

	// Hàm sử dụng để đổ ảnh vào bảng tính Excel
	// Kết quả trả về là dòng cuối cùng chứa hình ảnh sau khi tính toán, nếu lỗi sẽ trả về = 0
	// pSheet = Bảng tính sử dụng
	// szPathIMG = Đường dẫn hình ảnh sử dụng
	// iRowBegin, iColEnd = Vị trí chứa dòng & cột bắt đầu đổ hình ảnh
	// numAdd = Số lượng cột sẽ cộng thêm, tính từ vị trí chứa cột bắt đầu chèn ảnh (iColBegin)
	int _xlInsertPictureExcel(_WorksheetPtr pSheet, CString szPathIMG, int iRowBegin, int iColEnd, int numAdd);

	// Hàm tạo checkbox đặc biệt tại ô cell (nRow, nCol)
	// nRow, nCol = Vị trí sẽ tạo Checkbox
	// szNameRec = Tên rectange được tạo
	// szOnAction = Tên sự kiện được gán
	// szFormula = Công thức sử dụng để vào ô chứa checkbox
	// blFont = true --> Định dạng ô cell: Font = 'Wingdings 2'
	// nsize = Kích thước font chữ (14)
	bool _xlCreateCheckBox(_WorksheetPtr pSheet, int nRow, int nCol,
		CString szNameRec, CString szFormula = L"", CString szOnAction = L"",
		bool blFont = false, int nsize = 14);

	// Hàm định nghĩa NumberFormat ô cell
	bool _xlSetNumberFormat(RangePtr pCell, CString szFormat = L"General");

	// Hàm xác định ký tự ngăn cách (Chỉ chạy được bằng quyền Admin)
	// szDecimal = Ký tự ngăn cách hàng thập phân
	// szDigitGroup = Ký tự ngăn cách nhóm tiền tệ
	void _xlGetDecimalSymbol(CString &szDecimal, CString &szDigitGroup, int nGetRegistry = 1);

	// Hàm định nghĩa kiểu tiền tệ tài chính
	// szPositive = Kiểu số dương
	// szNegative = Kiểu số âm
	// szZero = Kiểu giá trị = 0
	// szText = Kiểu text
	// szDecimal = Ký tự ngăn cách hàng thập phân
	// szDigitGroup = Ký tự ngăn cách nhóm tiền tệ
	// blReplace = Tự động replace ký tự
	CString _xlAccountingFormat(
		CString szPositive = L"_(* #,##0.00_)", CString szNegative = L"_(* -#,##0.00_)",
		CString szZero = L"_(* \"0\"??_)", CString szText = L"_(@_)",
		CString szDecimal = L".", CString szDigitGroup = L",",
		bool blReplace = false);

	// Hàm khóa / mở khóa sheet
	// pSheet = Bảng tính cần khóa
	// szPassLock = Mật khẩu sử dụng
	// iLook = 1 -> Khóa | iLook = 0 -> Mở khóa
	int _xlLockSheet(_WorksheetPtr pSheet, CString szPassLock, int iLook);

	// Hàm khóa / mở khóa vùng dữ liệu Ecell
	// pSheet = Bảng tính cần khóa
	// pCell = Vùng dữ liệu cần khóa
	// szPassLock = Mật khẩu sử dụng
	// iLook = 1 -> Khóa | iLook = 0 -> Mở khóa
	int _xlLockCells(_WorksheetPtr pSheet, RangePtr pCell, CString szPassLock, int iLook);

	

/***********************************************************************************************/

	// Hàm kiểm tra phiên bản phần mềm
	int _CheckVersionSoftware();

	// Hàm chuyển từ kiểu Cstring --> char*
	char* _ConvertCstringToChars(CString szvalue);

	// Hàm chuyển từ kiểu char* --> Cstring
	CString _ConvertCharsToCstring(char* chr);

	// Hàm tách từ khóa chính thành các từ khóa con. Kết quả trả về là số lượng và chi tiết từ khóa con tìm được
	// vecPkey = vector lưu trữ danh sách từ khóa con
	// sKeyValue = Từ khóa chính
	// symbol1, symbol2 = Ký tự sử dụng để phân cách các từ khóa (symbol2 có thể bằng chính symbol1)
	// itypeFilter = 1 --> Lọc các từ khóa bị trùng lặp | itypeFilter = 0 --> Lấy tất cả từ khóa
	// itypeTrim = 1 --> Xóa khoảng trắng 2 đầu mỗi từ khóa | itypeFilter = 0 --> Giữ nguyên từ khóa không làm gì cả
	int _TackTukhoa(vector<CString> &vecPkey, CString sKeyValue, CString symbol1 = L";", CString symbol2 = L"|", int itypeFilter = 0, int itypeTrim = 1);

		// Hàm không chính xác khi thư mục được tạo có sử dụng tính năng chia sẻ (Dropbox, Google,...)
	// Hàm xác định tất cả các tệp tin trong thư mục được chỉ định
	// Kết quả trả về là tổng số files tìm thấy và lưu trữ thông tin vào '_finder' (_finder.GetFilePath(i).GetFileName()...)
	// pathFolder = Đường dẫn thư mục tìm kiếm
	// typeOfFile = Kiểu file cần tìm kiếm (xls,mdb,...)
	int _FindAllFile(CString pathFolder, CFileFinder &_finder, CString typeOfFile = L"*");

	// Hàm xác định tất cả các tệp tin trong thư mục được chỉ định
	// Kết quả trả về là tổng số files tìm thấy và lưu trữ thông tin vào 'vecPathFiles'
	// pathFolder = Đường dẫn thư mục tìm kiếm
	// vecPathFiles = Vector sử dụng để lưu trữ thông tin
	int _FindAllFileEx(CString pathFolder, vector<CString> &vecPathFiles);

	// Hàm xác định tất cả các thư mục trong thư mục được chỉ định
	// Kết quả trả về là tổng số thư mục tìm thấy và lưu trữ thông tin vào 'vecPathFolder'
	// szPathCha = Đường dẫn gốc sử dụng để tìm kiếm thư mục
	// vecPathFolder = Vector sử dụng để lưu trữ thông tin
	int _FindAllFolder(CString szPathCha, vector<CString> &vecPathFolder);

	// Hàm lấy đường dẫn của tệp tin 'fileName' được sử dụng
	CString _GetPathFolder(CString fileName);

	// Hàm ghi dữ liệu lên file unicode
	// vecTextLine = các giá trị lần lượt được ghi
	// spathFile = Đường dẫn file sử dụng (Nếu không có sẽ tự tạo file)
	// blClear = true --> Xóa toàn bộ dữ liệu trước khi ghi kết quả
	// itype = 1 --> Nếu không tồn tại file dừng lại luôn
	void _WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, bool blClear = false, int itype = 0);
	void _WriteMultiUnicode(CString strLine, CString spathFile, bool blClear = false, int itype = 0);

	// Hàm đọc dữ liệu unicode từ file text, kết quả trả về là tổng số dòng tìm thấy và lưu danh sách vào vector
	// spathFile = Đường dẫn tệp tin txt hoặc ini
	// vecTextLine = vector lưu trữ dữ liệu các dòng tìm được
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);

	// Hàm kiểm tra chuỗi có phải là số nguyên không?
	// Kết quả trả về TRUE nếu là số nguyên
	bool _IsNumber(CString &szNumber);
	
	// Hàm kiểm tra chuỗi có phải số La Mã không (Ví dụ: I, II,...)
	// Kết quả trả về TRUE nếu là số La Mã
	bool _IsNumberLaMa(CString szNumber);

	// Hàm kiểm tra chuỗi có phải số La Mã cấp 2 không (Ví dụ: I.1, I.2,...)
	// Kết quả trả về TRUE nếu là số La Mã cấp 2
	bool _IsNumberLaMaCap2(CString szNumber);

	// Hàm chuyển đổi số La Mã sang số nguyên dương, giới hạn từ 1-99
	int _ConvertLaMaToNumber(CString szLaMa);

	// Hàm chuyển đổi số nguyên dương sang số La Mã, giới hạn từ 1-99
	CString _ConvertNumberToLaMa(int number);

	// Hàm chuyển đổi sang chữ HOA
	CString _ConvertUpper(CString szNoidung);

	// Hàm chuyển đổi sang chữ thường
	CString _ConvertLower(CString szNoidung);

	// Hàm viết hoa chữ cái đầu tiên mỗi câu
	CString _ConvertUPProper(CString szNoidung);

	// Chỉ nên áp dụng trong trường hợp chuỗi có ký tự tiếng việt có dấu
	// Hàm Replace không phân biệt chữ HOA - thường
	CString _ReplaceMatchCase(CString szNoidung, CString szOld, CString szNew, bool blCase = true, bool blKDau = false);

	// Hàm chuyển đổi số sang chuỗi
	// number = Giá trị cần chuyển đổi sang chuỗi
	// iType (x) = 0..6 --> %0xd (001,002,...)
	CString _ConvertNumberToString(int number, int iType = 0);

	// Hàm convert không dấu
	// iLower = 1 --> Chuyển hết về chữ thường
	// iLower = 2 --> Chuyển hết về chữ HOA
	// iLower = 0 --> Bỏ qua không làm gì cả
	CString _ConvertKhongDau(CString szNoidung, int iLower = 0);

	// Hàm đổi giá trị từ double = ivalue (Bytes) sang KB,MB,GB,TB,...
	CString _ConvertBytes(double ivalue);

	// Hàm xác định thông tin file (bao gồm kích thước và ngày)
	// szDir = Đường dẫn thư mục cần tìm kiếm tệp tin
	// Giá trị trả về lưu trữ = szKichthuoc thước & szNgay
	void _GetFileAttributesEx(CString szDir, CString &szKichthuoc, CString &szNgay);

	// Hàm lấy giá trị hoặc thay đổi tên tiêu đề trên list: 'clist'
	// column = Cột cần thay đổi
	// szName = Giá trị sẽ thay thế
	// itype = 0 --> Lấy giá trị | itype = 1 --> Thay thế giá trị cũ
	CString _NameColumnHeading(CListCtrl &clist, int column, int itype = 1, CString szName = L"");
	
	// Hàm xác định windows 32 hay 64. Kết quả trả về True nếu là windows 64bit 
	bool _Is64BitWindows();

		// Hàm kiểm tra ứng dụng có chạy hay không (Ví dụ: GFilesManager.exe)
	bool _IsProcessRunning(const wchar_t *processName);

	// Hàm kiểm tra tệp tin hoặc thư mục (bao gồm cả Unicode) có tồn tại hay không (return = 1 --> tồn tại)
	// pathFile = đường dẫn tệp tin cần kiểm tra
	// itype = 1 thông báo nếu có lỗi xảy ra | itype = 0 bỏ qua thông báo
	int _FileExistsUnicode(CString pathFile, int itype = 0);

	// Hàm kiểm tra có phải thư mục không?
	bool _FolderExistsUnicode(CString dirName_in);

	// Convert chuỗi sang màu COLORREF
	// symbol = Ký tự ngăn cách giữa các giá trị RGB
	// Khi giá trị szColor = NULL: Nếu itype = 0 --> return = RGBWHITE | itype = 1 --> RGBBLACK;
	COLORREF _SetColorRGB(CString szColor, CString symbol = L",", int itype = 0);

	// Convert màu COLORREF sang chuỗi
	// symbol = Ký tự ngăn cách giữa các giá trị RGB
	// itype = 1 có thêm định dạng RGB(*,*,*)
	CString _GetColorRGB(COLORREF clrRGB, CString symbol = L",", int itype = 1);

	// Hàm gọi hộp thoại lấy màu
	// Trả về true nếu lấy được màu = clr & giá trị convert sang chuỗi 'szColor'
	bool _GetColorDlg(COLORREF &clr, CString &szColor);

	// Hàm lấy giá trị thứ (nIndex) từ cột thông số trả về kết quả lưu vào 'szReturn'
	// Nếu kết quả trả về false --> Không tồn tại giá trị nào ở cột thông số
	bool _GetValueThongso(CString szThongso, int nIndex, CString &szReturn);

	// Tooltip hướng dẫn
	// toolID = ID toolbox sử dụng
	// hDlg = Handle (mặc định = m_hWnd)
	// pszText = Nội dung hướng dẫn chi tiết
	HWND _CreateToolTip(int toolID, HWND hDlg, CString pszText);

	// Kiểm tra thanh Taskbar windows có bị ẩn không
	// Kết quả trả về iTaskHeight = -1 -- > Thanh Taskbar bị ẩn
	// Hoặc iTaskHeight =  0 - Bottom | 1 - Left | 2 - Top | 3 - Right
	int _IsTaskbarWndVisible(int &iTaskHeight);

	// Hàm format định dạng tiền VNĐ
	// szTien = Mệnh giá (giá trị là số)
	// szSteparator = Ký tự ngăn cách giữa hàng nghìn, trăm, ... 
	// szDecimal = Ký tự ngăn cách với hàng thập phân
	CString _FormatTienVND(CString szTien, CString szSteparator = L",", CString szDecimal = L".");

	// Hàm làm tròn số thực
	// doValue = Số thực cần làm tròn
	// nPrecision = Số thập phân sau dấu phẩy 
	double _RoundDouble(double doValue, int nPrecision);

	// Hàm lấy số ngẫu nhiên theo thời gian
	int _RandomTime();

	// Hàm lấy thời gian thực theo định dạng dd/MM/yyyy
	// ifulltime = 1 --> Trả vể đầy đủ cả ngày giờ dd/MM/yyyy HH:mm:ss
	CString _GetTimeNow(int ifulltime = 0);

	// Hàm đưa ngày tháng về định dạng chuẩn dd/MM/yyyy
	CString _FormatDate(CString szDate);

	// Hàm so sánh ngày tháng
	// date1, date2 : 2 ngày sử dụng để so sánh với nhau
	// Return = 0 --> date1 >  date2	(Ví dụ: date1 = 24/10/2018 & date2 = 12/08/2019)
	// Return = 1 --> date1 <= date2	(Ví dụ: date1 = 12/03/2018 & date2 = 27/01/2018)
	int _CompareDate(CString date1, CString date2);

	// Hàm cộng/ trừ ngày từ 1 ngày bất kỳ
	// szDate = ngày bất kỳ cần cộng
	// numAdd =  số ngày cần cộng/ trừ (numAdd > 0 : cộng | numAdd < 0 : trừ)
	CString	_DayNextPrev(CString szDate, int numAdd = 1);

	// Hàm bổ trợ cho hàm _NumDay2
	// Hàm convert ngày tháng ra giá trị số
	// day/month/year = ngày/tháng/năm
	int _DayValue(int day, int month, int year);

	// Hàm xác định khoảng cách giữa 2 ngày với nhau
	// dateMax = ngày lớn hơn, dateMin = ngày nhỏ hơn
	// cong = Giá trị cộng thêm nếu cần thiết
	// Return > 0 --> dateMax > dateMin
	int _NumDay2(CString dateMax, CString dateMin, int cong);

	// Kiểm tra ngày sử dụng là thứ mấy
	// szDate = Ngày cần kiểm tra
	CString _IsDayOfTheWeek(CString szDate);

	// Hàm xác định đường dẫn Desktop
	CString _GetDesktopDir();

	// Hàm gọi hàm từ DLL khác
	// szDllName = Tên DLL sử dụng
	// szFunction = Tên hàm trong DLL được gọi ra sử dụng
	bool _CallFunctionDLL(CString szDllName, CString szFunction, bool blQuit = true);

	// Hàm gọi hàm/thủ tục từ file XLA
	// szFileName = Tên file XLA sử dụng (Ví dụ: QLCL.xla)
	// szModule = Tên module trong XLA
	// szSub = Tên hàm/thủ tục trong XLA
	bool _CallModuleXLA(CString szFileName, CString szModule, CString szSub);

	// Hàm lựa chọn file trong thư mục
	// szDirFileSelect = Đường dẫn chứa tệp tin được lựa chọn 
	void _ShellExecuteSelectFile(CString szDirFileSelect);

	// Hàm tạo mới 1 thư mục được chỉ định
	// Có thể sử dụng để bổ trợ cho hàm _CreateDirectories
	BOOL _CreateNewFolder(CString directoryPath);

	// Hàm tạo nhiều thư mục con từ đường dẫn được chỉ định
	bool _CreateDirectories(CString pathName);

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

	// Hàm sao chép dữ liệu vào Clipboard (Có Unicode)
	// szNoidung = nội dung cần sao chép
	void _CopyClipboard(CString szNoidung);

	// Hàm paste dữ liệu từ Clipboard và trả về giá trị được sao chép (Có Unicode)
	CString _PasteClipboard();

	// Hàm sắp xếp nhanh mảng số
	// arr = Mảng số cần sắp xếp gồm N giá trị
	// l = vị trí đầu tiên trong mảng (=1)
	// r = vị trí cuối cùng (value = tổng số phần tử trong mảng = N)
	void _QuickSortArr(int *arr, int l, int r);

};

