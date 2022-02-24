#include "pch.h"
#include "base64.h"

static const string base64_chars =
"ABCDEFGHIJKLMNOPQRSTUVWXYZ"
"abcdefghijklmnopqrstuvwxyz"
"0123456789+/";

static inline bool is_base64(unsigned char c) {
	return (isalnum(c) || (c == '+') || (c == '/'));
}

Base64Ex::Base64Ex()
{
}

Base64Ex::~Base64Ex()
{
}

CString Base64Ex::ConvertstringtoUTF8(string ret)
{
	char *cstr = new char[ret.length() + 1];
	strcpy(cstr, ret.c_str());
	int num = ::MultiByteToWideChar(CP_UTF8, 0, cstr, -1, NULL, 0);
	wchar_t* wstr = new wchar_t[num + 1];
	::MultiByteToWideChar(CP_UTF8, 0, cstr, -1, wstr, num);
	wstr[num] = '\0';
	CString szval = L"";
	szval.Format(L"%s", wstr);
	return szval;
}

string Base64Ex::ConvertUTF8tostring(CString szval)
{
	wchar_t* wszString = new wchar_t[szval.GetLength() + 1];
	wcscpy(wszString, szval);
	int num = ::WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
	char* utf8 = new char[num + 1];
	::WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), utf8, num, NULL, NULL);
	utf8[num] = '\0';
	string str(utf8);
	return str;
}

string Base64Ex::base64_encode(unsigned char const* bytes_to_encode, unsigned int in_len)
{
	int i = 0, j = 0;
	unsigned char char_array_3[3];
	unsigned char char_array_4[4];
	string ret;

	while (in_len--)
	{
		char_array_3[i++] = *(bytes_to_encode++);
		if (i == 3)
		{
			char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
			char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
			char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
			char_array_4[3] = char_array_3[2] & 0x3f;

			for (i = 0; (i < 4); i++) ret += base64_chars[char_array_4[i]];
			i = 0;
		}
	}

	if (i)
	{
		for (j = i; j < 3; j++) char_array_3[j] = '\0';

		char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
		char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
		char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
		char_array_4[3] = char_array_3[2] & 0x3f;

		for (j = 0; (j < i + 1); j++) ret += base64_chars[char_array_4[j]];
		while ((i++ < 3)) ret += '=';
	}

	return ret;
}

string Base64Ex::base64_decode(string const& encoded_string)
{
	int in_len = encoded_string.size();
	int i = 0, j = 0, in_ = 0;
	unsigned char char_array_4[4], char_array_3[3];
	string ret;

	while (in_len-- && (encoded_string[in_] != '=') && is_base64(encoded_string[in_]))
	{
		char_array_4[i++] = encoded_string[in_];
		in_++;

		if (i == 4)
		{
			for (i = 0; i < 4; i++) char_array_4[i] = base64_chars.find(char_array_4[i]);

			char_array_3[0] = (char_array_4[0] << 2) + ((char_array_4[1] & 0x30) >> 4);
			char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
			char_array_3[2] = ((char_array_4[2] & 0x3) << 6) + char_array_4[3];

			for (i = 0; (i < 3); i++) ret += char_array_3[i];
			i = 0;
		}
	}

	if (i)
	{
		for (j = i; j < 4; j++) char_array_4[j] = 0;
		for (j = 0; j < 4; j++) char_array_4[j] = base64_chars.find(char_array_4[j]);

		char_array_3[0] = (char_array_4[0] << 2) + ((char_array_4[1] & 0x30) >> 4);
		char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
		char_array_3[2] = ((char_array_4[2] & 0x3) << 6) + char_array_4[3];

		for (j = 0; (j < i - 1); j++) ret += char_array_3[j];
	}

	return ret;
}

int Base64Ex::BaseRandomTime()
{
	time_t now = time(0);
	tm *ltm = localtime(&now);
	return (100 * (int)ltm->tm_min + (int)ltm->tm_sec + rand() % 1000);
}

CString Base64Ex::BaseEncodeEx(CString sValue, int itype)
{
	try
	{
		if (sValue == L"") return L"";
		if (itype == 1)
		{
			Function *ff = new Function;
			CString szChar = L"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

			for (int i = 0; i < 5; i++)
			{
				sValue += szChar.Mid((BaseRandomTime() + rand()) % 36, 1);
			}
		}
		else sValue += L"LEO89";

		CString szval = sValue;
		int iLen = szval.GetLength();
		int iDiv = (int)(iLen / 2);
		sValue = szval.Right(iLen - iDiv) + szval.Left(iDiv);

		string sencode = ConvertUTF8tostring(sValue);
		string en = base64_encode(reinterpret_cast<const unsigned char*>(sencode.c_str()), sencode.length());
		CString szbase(en.c_str());

		CString fx[15];
		iLen = szbase.GetLength();
		for (int i = 0; i < 6; i++) fx[i] = szbase.Mid(i, 1);
		szval.Format(L"%s%s%s%s%s%s%s", fx[1], fx[3], fx[5], fx[2], fx[0], fx[4], szbase.Right(iLen - 6));

		for (int i = 5; i >= 0; i--) fx[5 - i] = szval.Mid(iLen - i - 1, 1);
		sValue.Format(L"%s%s%s%s%s%s%s", szval.Left(iLen - 6), fx[3], fx[2], fx[0], fx[5], fx[1], fx[4]);

		szval = sValue;
		iDiv = (int)(iLen / 2);
		sValue = szval.Right(iLen - iDiv) + szval.Left(iDiv);

		string sencode2 = ConvertUTF8tostring(sValue);
		string en2 = base64_encode(reinterpret_cast<const unsigned char*>(sencode2.c_str()), sencode2.length());
		CString szbase2(en2.c_str());

		iLen = szbase2.GetLength();
		if (iLen < 16)
		{
			for (int i = 0; i < 8; i++) fx[i] = szbase2.Mid(i, 1);
			szval.Format(L"%s%s%s%s%s%s%s%s%s", fx[2], fx[0], fx[7], fx[1], fx[6], fx[4], fx[3], fx[5], szbase2.Right(iLen - 8));

			for (int i = 7; i >= 0; i--) fx[7 - i] = szval.Mid(iLen - i - 1, 1);
			sValue.Format(L"%s%s%s%s%s%s%s%s%s", szval.Left(iLen - 8), fx[1], fx[3], fx[6], fx[4], fx[7], fx[0], fx[5], fx[2]);
		}
		else
		{
			for (int i = 0; i < 12; i++) fx[i] = szbase2.Mid(i, 1);
			szval.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s", fx[9], fx[7], fx[0], fx[4], fx[11], fx[1], fx[2], fx[10], fx[3], fx[5], fx[6], fx[8], szbase2.Right(iLen - 12));

			for (int i = 11; i >= 0; i--) fx[11 - i] = szval.Mid(iLen - i - 1, 1);
			sValue.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s", szval.Left(iLen - 12), fx[8], fx[9], fx[6], fx[10], fx[5], fx[2], fx[0], fx[3], fx[11], fx[1], fx[7], fx[4]);
		}

		szval = sValue;
		iDiv = (int)(iLen / 2);
		sValue = szval.Right(iLen - iDiv) + szval.Left(iDiv);

		return sValue;
	}
	catch (...) {}
	return L"";
}

CString Base64Ex::BaseDecodeEx(CString sValue)
{
	try
	{
		if (sValue == L"") return L"";

		CString fx[15];
		CString szval = sValue;
		int iLen = sValue.GetLength();
		int iDiv = (int)(iLen / 2);
		sValue = szval.Right(iDiv) + szval.Left(iLen - iDiv);

		if (iLen < 16)
		{
			for (int i = 7; i >= 0; i--) fx[7 - i] = sValue.Mid(iLen - i - 1, 1);
			szval.Format(L"%s%s%s%s%s%s%s%s%s", sValue.Left(iLen - 8), fx[5], fx[0], fx[7], fx[1], fx[3], fx[6], fx[2], fx[4]);

			for (int i = 0; i < 8; i++) fx[i] = szval.Mid(i, 1);
			sValue.Format(L"%s%s%s%s%s%s%s%s%s", fx[1], fx[3], fx[0], fx[6], fx[5], fx[7], fx[4], fx[2], szval.Right(iLen - 8));
		}
		else
		{
			for (int i = 11; i >= 0; i--) fx[11 - i] = sValue.Mid(iLen - i - 1, 1);
			szval.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s", sValue.Left(iLen - 12), fx[6], fx[9], fx[5], fx[7], fx[11], fx[4], fx[2], fx[10], fx[0], fx[1], fx[3], fx[8]);

			for (int i = 0; i < 12; i++) fx[i] = szval.Mid(i, 1);
			sValue.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s", fx[2], fx[5], fx[6], fx[8], fx[3], fx[9], fx[10], fx[1], fx[11], fx[0], fx[7], fx[4], szval.Right(iLen - 12));
		}

		CT2CA pszConvert(sValue);
		const string sdecode = pszConvert;
		sValue = ConvertstringtoUTF8(base64_decode(sdecode));

		szval = sValue;
		iLen = szval.GetLength();
		iDiv = (int)(iLen / 2);
		sValue = szval.Right(iDiv) + szval.Left(iLen - iDiv);

		for (int i = 5; i >= 0; i--) fx[5 - i] = sValue.Mid(iLen - i - 1, 1);
		szval.Format(L"%s%s%s%s%s%s%s", sValue.Left(iLen - 6), fx[2], fx[4], fx[1], fx[0], fx[5], fx[3]);

		for (int i = 0; i < 6; i++) fx[i] = szval.Mid(i, 1);
		sValue.Format(L"%s%s%s%s%s%s%s", fx[4], fx[0], fx[3], fx[1], fx[5], fx[2], szval.Right(iLen - 6));

		CT2CA pszConvert2(sValue);
		const string sdecode2 = pszConvert2;
		sValue = ConvertstringtoUTF8(base64_decode(sdecode2));

		szval = sValue;
		iLen = szval.GetLength();
		iDiv = (int)(iLen / 2);
		sValue = szval.Right(iDiv) + szval.Left(iLen - iDiv);
		sValue = sValue.Left(iLen - 5);

		return sValue;
	}
	catch (...) {}
	return L"";
}