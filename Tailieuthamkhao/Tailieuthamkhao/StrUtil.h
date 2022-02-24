#if !defined(AFX_STRUTIL_H__5F96D586_85C7_494E_90A9_D54201F65F7C__INCLUDED_)
#define AFX_STRUTIL_H__5F96D586_85C7_494E_90A9_D54201F65F7C__INCLUDED_

#include <stdlib.h>
#include <string.h>
#include <string>
#include <vector>
#include <assert.h>
#include <time.h>

#ifdef WIN32
#include <windows.h>
#include <oleauto.h>
#endif

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class  __declspec(dllexport) StrUtil
{
	char *m_pData;
	void alloc(int len) {
		if (m_pData)
			m_pData = (char *)realloc(m_pData, len);
		else
			m_pData = (char *)malloc(len);
	}

	//	static wchar_t arrTCVN32Unicode[128];

	StrUtil(char *pData) {
		m_pData = pData;
	}

public:
	StrUtil() {
		m_pData = NULL;
	}

	char *getStrA()
	{
		return m_pData;
	}

	wchar_t *getStrW()
	{
		return (wchar_t *)m_pData;
	}

#ifdef _UNICODE
	char *getStr()
	{
		return m_pData;
	}
#else
	wchar_t *getStr()
	{
		return (wchar_t *)m_pData;
	}
#endif

	/*****************************************************************************************************************/

	template <class T>
	static unsigned int calcMaxLenT(const T *szFormat, va_list argList)
	{
#define FORCE_ANSI      0x10000
#define FORCE_UNICODE   0x20000
#define FORCE_INT64     0x40000
#ifndef max
#define max(x,y)	((x)>(y))?(x):(y)
#define delete_max
#endif
#ifndef min
#define min(x,y)	((x)<(y))?(x):(y)
#define delete_min
#endif
		const bool bIsUnicode = (sizeof(T) == 2);
		unsigned int uiMaxLen = 0;//in byte
		for (const T *lpsz = szFormat; *lpsz != 0; lpsz++)
		{
			// handle '%'(0x25) character, but watch out for '%%'
			if (*lpsz != 0x25 || *(++lpsz) == 0x25)
			{
				uiMaxLen++;
				continue;
			}

			unsigned int uiItemLen = 0;

			// handle '%'(0x25) character with format
			unsigned int uiWidth = 0;
			for (; *lpsz != 0; lpsz++)
			{
				// check for valid flags: #=0x23
				if (*lpsz == 0x23)
					uiMaxLen += 2;   // for '0x'
				else if (*lpsz == 0x2a)//*=0x2a
					uiWidth = va_arg(argList, int);
				else if (*lpsz == 0x2d || *lpsz == 0x2b || *lpsz == 0x30 ||
					*lpsz == 0x20)//-=2d, +=2b, 0=30 ' '=20
					;
				else // hit non-flag character
					break;
			}
			// get width and skip it
			if (uiWidth == 0)
			{
				// width indicated by
				while ((*lpsz >= 0x30) && (*lpsz <= 0x39) && (*lpsz != 0))
					uiWidth = uiWidth * 10 + *lpsz;
				//uiWidth = _ttoi(lpsz);
				//for (; *lpsz != 0 && _istdigit(*lpsz); lpsz = _tcsinc(lpsz))
				//	;
			}

			unsigned int uiPrecision = 0;
			if (*lpsz == 0x2e)//.=2e
			{
				// skip past '.' separator (width.precision)
				lpsz++;// = _tcsinc(lpsz);

				// get precision and skip it
				if (*lpsz == 0x2a)//*=2a
				{
					uiPrecision = va_arg(argList, int);
					lpsz++;// = _tcsinc(lpsz);
				}
				else
				{
					while ((*lpsz >= 0x30) && (*lpsz <= 0x39) && (*lpsz != 0))
						uiPrecision = uiPrecision * 10 + *lpsz++ - 0x30;
				}
			}

			// should be on type modifier or specifier
			int nModifier = 0;
			switch (*lpsz)
			{
				// modifiers that affect size
			case 0x68://'h':
				nModifier = FORCE_ANSI;
				lpsz++;// = _tcsinc(lpsz);
				break;
			case 0x6c://'l':
				nModifier = FORCE_UNICODE;
				lpsz++;// = _tcsinc(lpsz);
				break;
			case 0x49://'I':
				if (*(lpsz + 1) == 0x36 && *(lpsz + 2) == 0x34)
				{
					lpsz += 3;
					nModifier = FORCE_INT64;
				}
				break;

				// modifiers that do not affect size
			case 0x46://'F':
			case 0x4e://'N':
			case 0x4c://'L':
				lpsz++;// = _tcsinc(lpsz);
				break;
			}

			// now should be on specifier
			switch (*lpsz | nModifier)
			{
				// single characters
			case 0x63://'c':
			case 0x43://'C':
				uiItemLen = 2;
				va_arg(argList, T);
				break;
			case 0x63 | FORCE_ANSI:
			case 0x43 | FORCE_ANSI:
				uiItemLen = 2;
				va_arg(argList, char);
				break;
			case 0x63 | FORCE_UNICODE:
			case 0x43 | FORCE_UNICODE:
				uiItemLen = 2;
				va_arg(argList, wchar_t);
				break;

				// strings
			case 0x73://'s':
			{
				const T *pstrNextArg = va_arg(argList, const T *);
				if (pstrNextArg == NULL)
					uiItemLen = 6;  // "(null)"
				else
				{
					//while(*pstrNextArg++ != 0)
					//	uiItemLen++;
					if (bIsUnicode)
						uiItemLen = wcslen((const wchar_t*)pstrNextArg);
					else
						uiItemLen = strlen((const char*)pstrNextArg);
					uiItemLen = max(1, uiItemLen);
				}
			}
			break;

			case 0x53://'S':
				if (bIsUnicode)
				{
					char *pstrNextArg = va_arg(argList, char *);
					if (pstrNextArg == NULL)
						uiItemLen = 6; // "(null)"
					else
					{
						uiItemLen = strlen(pstrNextArg);
						uiItemLen = max(1, uiItemLen);
					}
				}
				else
				{
					wchar_t *pstrNextArg = va_arg(argList, wchar_t *);
					if (pstrNextArg == NULL)
						uiItemLen = 6;  // "(null)"
					else
					{
						uiItemLen = wcslen(pstrNextArg);
						uiItemLen = max(1, uiItemLen);
					}
				}
				break;

			case 0x73 | FORCE_ANSI:
			case 0x53 | FORCE_ANSI:
			{
				char *pstrNextArg = va_arg(argList, char *);
				if (pstrNextArg == NULL)
					uiItemLen = 6; // "(null)"
				else
				{
					uiItemLen = strlen(pstrNextArg);
					uiItemLen = max(1, uiItemLen);
				}
			}
			break;

			case 0x73 | FORCE_UNICODE:
			case 0x53 | FORCE_UNICODE:
			{
				wchar_t *pstrNextArg = va_arg(argList, wchar_t *);
				if (pstrNextArg == NULL)
					uiItemLen = 6; // "(null)"
				else
				{
					uiItemLen = wcslen(pstrNextArg);
					uiItemLen = max(1, uiItemLen);
				}
			}
			break;
			}

			// adjust nItemLen for strings
			if (uiItemLen != 0)
			{
				if (uiPrecision != 0)
					uiItemLen = (uiItemLen > uiPrecision) ? uiPrecision : uiItemLen;
				uiItemLen = (uiItemLen > uiWidth) ? uiItemLen : uiWidth;
			}
			else
			{
				switch (*lpsz)
				{
					// integers
				case 'd':
				case 'i':
				case 'u':
				case 'x':
				case 'X':
				case 'o':
					if (nModifier & FORCE_INT64)
#ifdef WIN32
						va_arg(argList, __int64);
#else
						va_arg(argList, long long);
#endif //WIN32
					else
						va_arg(argList, int);
					uiItemLen = 32;
					uiItemLen = max(uiItemLen, uiWidth + uiPrecision);
					break;

				case 'e':
				case 'g':
				case 'G':
					va_arg(argList, double);
					uiItemLen = 128;
					uiItemLen = max(uiItemLen, uiWidth + uiPrecision);
					break;

				case 'f':
				{
					va_arg(argList, double);
					// 312 == strlen("-1+(309 zeroes).")
					// 309 zeroes == max precision of a double
					// 6 == adjustment in case precision is not specified,
					//   which means that the precision defaults to 6
					uiItemLen = max(uiWidth, 312 + uiPrecision + 6);
				}
				break;

				case 'p':
					va_arg(argList, void*);
					uiItemLen = 32;
					uiItemLen = max(uiItemLen, uiWidth + uiPrecision);
					break;

					// no output
				case 'n':
					va_arg(argList, int*);
					break;

				default:
					assert(0);  // unknown formatting option
				}
			}

			// adjust nMaxLen for output nItemLen
			uiMaxLen += uiItemLen;
		}
#undef FORCE_ANSI
#undef FORCE_UNICODE
#undef FORCE_INT64
#ifdef delete_min
#undef min
#endif
#ifdef delete_max
#undef max
#endif
		va_end(argList);
		return uiMaxLen;
	}

	/*****************************************************************************************************************/

	virtual ~StrUtil() {
		if (m_pData)
			free(m_pData);
		//delete []m_pData;
	}

	/*****************************************************************************************************************/

	inline wchar_t *ansi2unicode(const char *pStr, int len = -1) {
		if (len < 0)len = strlen(pStr);
		int _len = mbstowcs(NULL, pStr, len);
		if (_len < 0)return NULL;
		alloc(sizeof(wchar_t)*(_len + 1));
		_len = mbstowcs((wchar_t *)m_pData, pStr, len);
		if (_len < 0)
			return NULL;
		((wchar_t *)m_pData)[_len] = 0;
		return (wchar_t *)m_pData;
	}

	/*****************************************************************************************************************/

	inline char *unicode2ansi(const wchar_t *pStr, int len = -1) {
		if (len < 0)len = wcslen(pStr);
		int _len = wcstombs(NULL, pStr, len);
		if (_len < 0)return NULL;
		alloc(_len + 1);
		_len = wcstombs(m_pData, pStr, len);
		if (_len < 0)
			return NULL;
		m_pData[_len] = 0;
		return m_pData;
	}

	/*****************************************************************************************************************/

#ifdef _UNICODE
	inline const wchar_t *unicode_ansi_convert(const char *pStr, int len = -1) {
		return ansi2unicode(pStr, len);
	}
	inline const wchar_t *unicode_ansi_convert(const wchar_t *pStr, int len = -1) {
		return pStr;
	}
#else
	inline const char *unicode_ansi_convert(const char *pStr, int len = -1) {
		return pStr;
	}
	inline const char *unicode_ansi_convert(const wchar_t *pStr, int len = -1) {
		return unicode2ansi(pStr, len);
	}
#endif

	/*****************************************************************************************************************/

	inline const char *toAnsi(const char *pStr, int len = -1) {
		return pStr;
	}

	/*****************************************************************************************************************/

	inline const char *toAnsi(const wchar_t *pStr, int len = -1) {
		return unicode2ansi(pStr, len);
	}

	/*****************************************************************************************************************/

	inline const wchar_t *toUnicode(const char *pStr, int len = -1) {
		return ansi2unicode(pStr, len);
	}

	/*****************************************************************************************************************/

	inline const wchar_t *toUnicode(const wchar_t *pStr, int len = -1) {
		return pStr;
	}

	/*****************************************************************************************************************/

	template<class _IN, class _OUT>
	inline static void split(const _IN *pIn, _IN split, std::vector<_OUT> &out)
	{
		_IN buf[1024];
		_IN *pe = buf;
		while (true)
		{
			if (*pIn == split || *pIn == 0)
			{
				*pe = 0;
				out.push_back(buf);
				pe = buf;
				if (*pIn == 0)
					return;
			}
			else
			{
				*pe = *pIn;
				pe++;
			}
			pIn++;
		}
	}

	/*****************************************************************************************************************/

	template<class _IN, class _OUT>
	inline static void split(const _IN *pIn, const _IN *split, std::vector<_OUT> &out)
	{
		_IN buf[1024];
		_IN *pe = buf;
		while (true)
		{
			if (*pIn == 0 || 0 != strchr(split, *pIn))
			{
				*pe = 0;
				out.push_back(buf);
				pe = buf;
				if (*pIn == 0)
					return;
			}
			else
			{
				*pe = *pIn;
				pe++;
			}
			pIn++;
		}
	}

	/*****************************************************************************************************************/

	//	template<class _IN, class _OUT>
	//	inline static void split(const _IN *pIn, _IN split, _OUT &out, int &size)
	//	{
	//		_IN buf[64];
	//		_IN *pe = buf;
	//		size = 0;
	//		while(true)
	//		{
	//			if(*pIn == split || *pIn == 0)
	//			{
	//				*pe = 0;
	//				out << buf;
	//				//out->push_back(buf);
	//				pe = buf;
	//				size++;
	//				if(*pIn == 0)
	//					return;
	//			}else
	//			{
	//				*pe = *pIn;
	//				pe++;
	//			}
	//			pIn++;
	//		}
	//	}

	/*****************************************************************************************************************/
	char *toAnsiTCVN3(wchar_t *pstr, int len = -1)
	{
		if (0 == pstr)return 0;
		if (len < 0)len = wcslen(pstr);
		alloc(len + 1);
		char *c = m_pData;
		for (; *pstr != 0; pstr++, c++)
		{
			switch (*pstr) {
			case 193:
			case 0xE1:*c = '¸'; break;//AS
			case 192:
			case 0xE0:	*c = 'µ';	break;//AF
			case 7842:
			case 0x1EA3:*c = '¶';	break;//AR
			case 7840:
			case 0x1EA1:*c = '¹';	break;//AJ
			case 195:case 0xE3:	*c = '·';	break;//AX
			case 0x0103:*c = '¨';	break;
			case 0x1EAF:*c = '¾';	break;
			case 0x1EB1:*c = '»';	break;
			case 0x1EB3:*c = '¼';	break;
			case 0x1EB7:*c = 'Æ';	break;
			case 0x1EB5:*c = '½';	break;
			case 0xE2:	*c = '©';	break;//AA
			case 7844:case 0x1EA5:*c = 'Ê';	break;//AAS
			case 7846:case 0x1EA7:*c = 'Ç';	break;//AAF
			case 7848:case 0x1EA9:*c = 'È';	break;//AAR
			case 7852:case 0x1EAD:*c = 'Ë';	break;//AAJ
			case 7850:case 0x1EAB:*c = 'É';	break;//AAX
			case 0xE9:	*c = 'Ð';	break;
			case 0xE8:	*c = 'Ì';	break;
			case 0x1EBB:*c = 'Î';	break;
			case 0x1EB9:*c = 'Ñ';	break;
			case 0x1EBD:*c = 'Ï';	break;
			case 0xEA:	*c = 'ª';	break;
			case 0x1EBF:*c = 'Õ';	break;
			case 0x1EC1:*c = 'Ò';	break;
			case 0x1EC3:*c = 'Ó';	break;
			case 0x1EC7:*c = 'Ö';	break;
			case 0x1EC5:*c = 'Ô';	break;
			case 0xED:	*c = 'Ý';	break;
			case 0xEC:	*c = '×';	break;
			case 0x1EC9:*c = 'Ø';	break;
			case 0x1ECB:*c = 'Þ';	break;
			case 0x0129:*c = 'Ü';	break;
			case 0xF3:	*c = 'ã';	break;
			case 0xF2:	*c = 'ß';	break;
			case 0x1ECF:*c = 'á';	break;
			case 0x1ECD:*c = 'ä';	break;
			case 0xF5:	*c = 'â';	break;
			case 0xF4:	*c = '«';	break;
			case 0x1ED1:*c = 'è';	break;
			case 0x1ED3:*c = 'å';	break;
			case 0x1ED5:*c = 'æ';	break;
			case 0x1ED9:*c = 'é';	break;
			case 0x1ED7:*c = 'ç';	break;
			case 0x01A1:*c = '¬';	break;
			case 0x1EDB:*c = 'í';	break;
			case 0x1EDD:*c = 'ê';	break;
			case 0x1EDF:*c = 'ë';	break;
			case 0x1EE3:*c = 'î';	break;
			case 0x1EE1:*c = 'ì';	break;
			case 0xFA:	*c = 'ó';	break;
			case 0xF9:	*c = 'ï';	break;
			case 0x1EE7:*c = 'ñ';	break;
			case 0x1EE5:*c = 'ô';	break;
			case 0x0169:*c = 'ò';	break;
			case 0x01B0:*c = '­';	break;
			case 0x1EE9:*c = 'ø';	break;
			case 0x1EEB:*c = 'õ';	break;
			case 0x1EED:*c = 'ö';	break;
			case 0x1EF1:*c = 'ù';	break;
			case 0x1EEF:*c = '÷';	break;
			case 253:	*c = 'ý';	break;
			case 0x1EF3:*c = 'ú';	break;
			case 0x1EF7:*c = 'û';	break;
			case 0x1EF5:*c = 'þ';	break;
			case 0x1EF9:*c = 'ü';	break;
			case 0x0111:*c = '®';	break;
			case 0x0110:*c = '§';	break;
			case 0x0102:*c = '¡';	break;
			case 0xC2:	*c = '¢';	break;
			case 0xCA:	*c = '£';	break;
			case 0xD4:	*c = '¤';	break;
			case 0x01A0:*c = '¥';	break;
			case 0x01AF:*c = '¦';	break;
			case 7888: *c = 'è'; break;
			default:
				*c = (char)*pstr;
			};
		}
		*c = 0;
		return m_pData;
	}

	/*****************************************************************************************************************/

	wchar_t *fromAnsiTCVN3(const char *szTcvn3, int len = -1)
	{
		if (szTcvn3 == 0) return 0;

		static wchar_t arrTCVN32Unicode[128] = {
			//from 128 to 255
			128,			129,			130,			131,
			132,			133,			134,			135,
			136,			137,			138,			139,
			140,			141,			142,			143,
			144,			145,			146,			147,
			148,			149,			150,			151,
			152,			153,			154,			155,
			156,			157,			158,			159,
			160,			0x0102,/*¡*/	162,			0xCA,/*£*/
			0xD4,/*¤*/		0x01A0,/*¥*/	0x01AF,/*¦*/	0x0110,/*§*/
			0x0103,/*¨*/	0xE2,/*©*/		0xEA,/*ª*/		0xF4,/*«*/
			0x01A1,/*¬*/	0x01B0,/*­*/	0x0111,/*®*/	175,
			176,			177,			178,			179,
			180,			0xE0,/*µ*/		0x1EA3,/*¶*/	0xE3,/*·*/
			0xE1,/*¸*/		0x1EA1,/*¹*/	186,			0x1EB1,/*»*/
			0x1EB3,/*¼*/	0x1EB5,/*½*/	0x1EAF,/*¾*/	191,
			192,			193,			0xC2,/*¢*/			195,
			196,			197,			0x1EB7,/*Æ*/	0x1EA7,/*Ç*/
			0x1EA9,/*È*/	0x1EAB,/*É*/	0x1EA5,/*Ê*/	0x1EAD,/*Ë*/
			0xE8,/*Ì*/		205,			0x1EBB,/*Î*/	0x1EBD,/*Ï*/
			0xE9,/*Ð*/		0x1EB9,/*Ñ*/	0x1EC1,/*Ò*/	0x1EC3,/*Ó*/
			0x1EC5,/*Ô*/	0x1EBF,/*Õ*/	0x1EC7,/*Ö*/	0xEC,/*×*/
			0x1EC9,/*Ø*/	217,			218,			219,
			0x0129,/*Ü*/	0xED,/*Ý*/		0x1ECB,/*Þ*/	0xF2,/*ß*/
			224,			0x1ECF,/*á*/	0xF5,/*â*/		0xF3,/*ã*/
			0x1ECD,/*ä*/	0x1ED3,/*å*/	0x1ED5,/*æ*/	0x1ED7,/*ç*/
			0x1ED1,/*è*/	0x1ED9,/*é*/	0x1EDD,/*ê*/	0x1EDF,/*ë*/
			0x1EE1,/*ì*/	0x1EDB,/*í*/	0x1EE3,/*î*/	0xF9,/*ï*/
			240,			0x1EE7,/*ñ*/	0x0169,/*ò*/	0xFA,/*ó*/
			0x1EE5,/*ô*/	0x1EEB,/*õ*/	0x1EED,/*ö*/	0x1EEF,/*÷*/
			0x1EE9,/*ø*/	0x1EF1,/*ù*/	0x1EF3,/*ú*/	0x1EF7,/*û*/
			0x1EF9,/*ü*/	253,/*ý*/		0x1EF5,/*þ*/	255
		};

		if (len < 0)len = strlen(szTcvn3);
		alloc(sizeof(wchar_t)*(len + 1));

		register unsigned char c;
		wchar_t * p = (wchar_t *)m_pData;
		for (c = *szTcvn3; (c = *szTcvn3) != 0; szTcvn3++, p++) {
			if (c == 1) {
				*p = *(wchar_t *)(szTcvn3 + 1);
				szTcvn3 += 2;
			}
			else if (c == 2) {
				*p = szTcvn3[1];
				szTcvn3++;
			}
			else if (c < 128) *p = c;
			else *p = arrTCVN32Unicode[c - 128];
		};
		*p = 0;
		return (wchar_t *)m_pData;
	}

	/*****************************************************************************************************************/

	template<class TYPE>
	inline static TYPE *toLower(TYPE *str) {
		TYPE *tg = str;
		while (0 != (*tg))
		{
			if ('A' <= *tg && 'Z' >= *tg)
				*tg |= 32;
			tg++;
		}
		return str;
	}

	/*****************************************************************************************************************/

	template<class TYPE>
	inline static TYPE *toUpper(TYPE *str) {
		TYPE *tg = str;
		while (*tg != 0)
		{
			if ('A' <= *tg && 'Z' >= *tg)
				str[i] &= 0xFFFFFFDF;
			tg++;
		}
		return str;
	}

	/*****************************************************************************************************************/

	template<class TYPE>
	inline static TYPE *trimLeft(TYPE *str)
	{
		if (0 == str || 0 == *str)return str;
		TYPE *p = str;
		while ((*p != 0) && (*p == ' ' || *p == '\t' || *p == 10 || *p == 13))
			p++;
		return p;
	}

	/*****************************************************************************************************************/

	template<class TYPE>
	inline static TYPE *trimRight(TYPE *str, int len)
	{
		if (0 == str || 0 == *str)return str;
		TYPE *p = str + len - 1;
		while ((p != str) && (*p == ' ' || *p == '\t' || *p == 10 || *p == 13))
			p--;
		if (p == str && len > 1)
			p[0] = 0;
		else
			p[1] = 0;
		return str;
	}

	/*****************************************************************************************************************/

	static int length(const char *str) {
		if (0 == str)return 0;
		return strlen(str);
	}

	/*****************************************************************************************************************/

	static int length(const wchar_t *str) {
		if (0 == str)return 0;
		return wcslen(str);
	}

	/*****************************************************************************************************************/

	static const char *strchr(const char *str, int c) {
		if (0 == str)return 0;
		return ::strchr(str, c);
	}

	/*****************************************************************************************************************/

	static const wchar_t *strchr(const wchar_t *str, int c) {
		if (0 == str)return 0;
		return ::wcschr(str, c);
	}

	/*****************************************************************************************************************/

	static const char *strstr(const char *str, const char *strCharSet) {
		if (0 == str)return 0;
		return ::strstr(str, strCharSet);
	}

	/*****************************************************************************************************************/

	static const wchar_t *strstr(const wchar_t *str, const wchar_t *strCharSet) {
		if (0 == str)return 0;
		return ::wcsstr(str, strCharSet);
	}

	/*****************************************************************************************************************/

	static int strcmp(const char *str, const char *str2) {
		return ::strcmp(str, str2);
	}

	/*****************************************************************************************************************/

	static int strcmp(const wchar_t *str, const wchar_t *str2) {
		return ::wcscmp(str, str2);
	}

	/*****************************************************************************************************************/

	static int stricmp(const char *str, const char *str2) {
		return ::stricmp(str, str2);
	}

	/*****************************************************************************************************************/

	static int stricmp(const wchar_t *str, const wchar_t *str2) {
		return ::wcsicmp(str, str2);
	}

	/*****************************************************************************************************************/

	static int strnicmp(const char *str, const char *str2, int n) {
		return ::strnicmp(str, str2, n);
	}

	/*****************************************************************************************************************/

	static int strnicmp(const wchar_t *str, const wchar_t *str2, int n) {
		return ::wcsnicmp(str, str2, n);
	}

	/*****************************************************************************************************************/

	static double toDouble(const char *str) {
		if (0 == str)return 0;
		char *end;
		return strtod(str, &end);
	}

	/*****************************************************************************************************************/

	static double toDouble(const wchar_t *str) {
		if (0 == str)return 0;
		wchar_t *end;
		return wcstod(str, &end);
	}

	/*****************************************************************************************************************/

	static long toLong(const char *str) {
		if (0 == str)return 0;
		char *end;
		return strtol(str, &end, 10);
	}

	/*****************************************************************************************************************/

	static long toLong(const wchar_t *str) {
		if (0 == str)return 0;
		wchar_t *end;
		return wcstol(str, &end, 10);
	}

	/*****************************************************************************************************************/

	static long toULong(const char *str) {
		if (0 == str)return 0;
		char *end;
		return strtoul(str, &end, 10);
	}

	/*****************************************************************************************************************/

	static long toULong(const wchar_t *str) {
		if (0 == str)return 0;
		wchar_t *end;
		return wcstoul(str, &end, 10);
	}

	/*****************************************************************************************************************/

	template<class T>
	inline static T replaceCPrefix(const T &str, int *pnumreplace = 0)
	{
		if (str.length() <= 0)
			return T();

		T ret;
		const T::value_type *run = str.c_str();
		if (pnumreplace)
			*pnumreplace = 0;
		while (0 != *run)
		{
			switch (*run)
			{
			case '\\':
				run++;
				switch (*run)
				{
				case '\\':	ret += *run;				break;
				case 'n':	ret += T::value_type('\n');	break;
				case 't':	ret += T::value_type('\t');	break;
				case 'r':	ret += T::value_type('\r');	break;
				case 'b':	ret += T::value_type('\b');	break;
				}
				if (pnumreplace)
					(*pnumreplace)++;
				break;
			default:
				ret += *run;
			}
			run++;
		}
		return ret;
	}
	/*****************************************************************************************************************/

	template<class T>
	inline static T removeQuote(const T &str)
	{
		if (str.length() <= 0)
			return T();

		return str.substr(1, str.length() - 2);
	}

	/*****************************************************************************************************************/

	char *fromIntA(int val)
	{
		alloc(32);
		return ltoa(val, m_pData, 10);
	}

	/*****************************************************************************************************************/

	wchar_t *fromIntU(int val)
	{
		alloc(64);
		return _ltow(val, (wchar_t *)m_pData, 10);
	}

	/*****************************************************************************************************************/

#ifdef _UNICODE
	wchar_t *fromInt(int val)
	{
		return fromIntU(val);
	}
#else
	char *fromInt(int val)
	{
		return fromIntA(val);
	}
#endif

	/*****************************************************************************************************************/

	char *format(const char *szFormat, ...)
	{
		va_list argList;
		va_start(argList, szFormat);
		unsigned int uiLen = calcMaxLenT(szFormat, argList);
		alloc(uiLen + 1);
		vsprintf(m_pData, szFormat, argList);
		return m_pData;
	}

	/*****************************************************************************************************************/

	wchar_t *format(const wchar_t *szFormat, ...)
	{
		va_list argList;
		va_start(argList, szFormat);
		unsigned int uiLen = calcMaxLenT(szFormat, argList);
		alloc((uiLen + 1) << 1);
		vswprintf((wchar_t *)m_pData, szFormat, argList);
		return (wchar_t *)m_pData;
	}

	/*****************************************************************************************************************/

	template<class CHAR_TYPE>
	static CHAR_TYPE *toLowerTCVN3(CHAR_TYPE *pstr)
	{
		if (0 == pstr)return 0;
		for (; *pstr != 0; pstr++)
		{
			switch (*pstr) {
			case '§': *pstr = '®';	break;
			case '¡': *pstr = '¨';	break;
			case '¢': *pstr = '©';	break;
			case '£': *pstr = 'ª';	break;
			case '¤': *pstr = '«';	break;
			case '¥': *pstr = '¬';	break;
			case '¦': *pstr = '­';	break;
			default:
				*pstr = towlower(*pstr);
			};
		}
		return pstr;
	}

	/*****************************************************************************************************************/

#ifdef WIN32
	inline char *toUtf8(const wchar_t *szU, int len = -1, bool convert = false)
	{
		if (len <= 0)
			len = wcslen(szU);
		std::wstring tg(szU, len);
		int pos = 0;
		if (convert)
		{

			while (true)
			{
				pos = tg.find_first_of(L"&\"<>'", pos);
				if (pos == tg.npos)
					break;
				switch (tg[pos])
				{
				case L'&':
					tg.replace(pos, 1, L"&amp;");
					pos += 5;
					break;
				case L'"':
					tg.replace(pos, 1, L"&quot;");
					pos += 6;
					break;
				case L'<':
					tg.replace(pos, 1, L"&lt;");
					pos += 4;
					break;
				case L'>':
					tg.replace(pos, 1, L"&gt;");
					pos += 4;
					break;
				case L'\'':
					tg.replace(pos, 1, L"&apos;");
					pos += 4;
					break;
				default:
					pos++;
				}
			}
		}
		pos = WideCharToMultiByte(CP_UTF8, 0, tg.c_str(), tg.length(), 0, 0, NULL, NULL);
		alloc(pos + 1);
		pos = WideCharToMultiByte(CP_UTF8, 0, tg.c_str(), tg.length(), m_pData, pos, NULL, NULL);
		m_pData[pos] = 0;
		return m_pData;
	}

	/*****************************************************************************************************************/

	inline wchar_t *fromUtf8(const char *szUtf8, int len = -1)
	{
		if (len <= 0)
			len = strlen(szUtf8);
		int buf_size = MultiByteToWideChar(CP_UTF8, 0, szUtf8, len, 0, 0) + 2;

		alloc(buf_size * 2);

		buf_size = MultiByteToWideChar(CP_UTF8, 0, szUtf8, len, (wchar_t *)m_pData, buf_size);
		((wchar_t *)m_pData)[buf_size] = 0;
		return (wchar_t *)m_pData;
	}

#endif

	/*****************************************************************************************************************/

		//inline static char *strpbrk(const char *str, const char *charSet)
		//{
		//	if(0 == str)return 0;
		//	return strpbrk(str, charSet);
		//}

	/*****************************************************************************************************************/

		//inline static wchar_t *strpbrk(const wchar_t *str, const wchar_t *charSet)
		//{
		//	if(0 == str)return 0;
		//	return wcspbrk(str, charSet);
		//}

	/*****************************************************************************************************************/

	template<class T>
	inline const T *replaceXMLTagT(const T *szStr, unsigned int iLen, unsigned int *pLen = 0)
	{
		const T *p = szStr;
		T *q;
		int iCount = 0;
		if (iLen < 0)
			iLen = length(szStr);
		/* count number of change */
		while (*p)
		{
			if (*p == '&' || *p == '"' || *p == '<' || *p == '>' || *p == '\'')iCount++;
			p++;
		}
		/* change if iCount > 0 */
		if (iCount)
		{
			alloc((iLen + 5 * iCount) * sizeof(T));
			p = szStr;
			q = (T *)m_pData;
			do
			{
				switch (*p)
				{
				case '&':	// -> "&amp;"
					*q++ = '&';				*q++ = 'a';				*q++ = 'm';
					*q++ = 'p';				*q++ = ';';				break;
				case '"':	// -> "&quot;"
					*q++ = '&';				*q++ = 'q';				*q++ = 'u';
					*q++ = 'o';				*q++ = 't';				*q++ = ';';				break;
				case '<':	// -> "&lt;"
					*q++ = '&';				*q++ = 'l';				*q++ = 't';
					*q++ = ';';				break;
				case '>':	// -> "&gt;"
					*q++ = '&';				*q++ = 'g';				*q++ = 't';
					*q++ = ';';				break;
				case '\'':	// -> "&apos;"
					*q++ = '&';				*q++ = 'a';				*q++ = 'p';
					*q++ = 'o';				*q++ = 's';				*q++ = ';';				break;
				default:
					*q++ = *p;

				}
				p++;
			} while (*p);
			*q = 0;
			if (pLen)*pLen = q - (T *)m_pData;
			return (T *)m_pData;
		}
		if (pLen)*pLen = iLen;
		return szStr;
	}

	/*****************************************************************************************************************/

	template<class T>
	inline const T *restoreXMLTagT(const T *szStr, unsigned int iLen = -1, unsigned int *pLen = 0)
	{
		const T *p = szStr;
		T *q;
		int iCount = 0;
		int iRemain;
		if (iLen < 0)
			iLen = length(szStr);
		/* count number of change */
		while (*p)
		{
			if (*p == '&')
			{
				iRemain = iLen - (p - szStr);
				if (iRemain >= 4)
				{
					if ((*(p + 1) == 'l' || *(p + 1) == 'g') && *(p + 2) == 't' && *(p + 3) == ';')
					{
						iCount++;
						p += 3;
					}
					else if (iRemain >= 5)
					{
						if (*(p + 1) == 'a' && *(p + 2) == 'm' && *(p + 3) == 'p' && *(p + 4) == ';')
						{
							iCount++;
							p += 4;
						}
						else if (iRemain >= 6)
						{
							if ((*(p + 1) == 'a' || *(p + 1) == 'q') && (*(p + 2) == 'u' || *(p + 2) == 'p') &&
								*(p + 3) == 'o' && (*(p + 4) == 't' || *(p + 4) == 's') && *(p + 5) == ';')
							{
								iCount++;
								p += 5;
							}
						}
					}
				}
			}
			p++;
		}
		/* change if iCount > 0 */
		if (iCount)
		{
			alloc(iLen * sizeof(T));
			p = szStr;
			q = (T *)m_pData;

			while (*p)
			{
				if (*p == '&')
				{
					iRemain = iLen - (p - szStr);
					if (iRemain >= 4)
					{
						if ((*(p + 1) == 'l' || *(p + 1) == 'g') && *(p + 2) == 't' && *(p + 3) == ';')
						{
							*q++ = (*(p + 1) == 'l') ? '<' : '>';
							p += 4;
							continue;
						}
						else if (iRemain >= 5)
						{
							if (*(p + 1) == 'a' && *(p + 2) == 'm' && *(p + 3) == 'p' && *(p + 4) == ';')
							{
								*q++ = '&';
								p += 5;
								continue;
							}
							else if (iRemain >= 6)
							{
								if ((*(p + 1) == 'a' || *(p + 1) == 'q') && (*(p + 2) == 'u' || *(p + 2) == 'p') &&
									*(p + 3) == 'o' && (*(p + 4) == 't' || *(p + 4) == 's') && *(p + 5) == ';')
								{
									*q++ = (*(p + 1) == 'a') ? '\'' : '"';
									p += 6;
									continue;
								}
							}
						}
					}
				}
				*q++ = *p++;
			}

			*q = 0;
			if (pLen)*pLen = q - (T *)m_pData;
			return (T *)m_pData;
		}
		if (pLen)*pLen = iLen;
		return szStr;
	}

	/*****************************************************************************************************************/

	inline char *fromDateTime(const char *format, const double &dt)
	{
		UDATE ud;
		if (S_OK != VarUdateFromDate(dt, 0, &ud))
			return 0;
		struct tm tmTemp;
		tmTemp.tm_sec = ud.st.wSecond;
		tmTemp.tm_min = ud.st.wMinute;
		tmTemp.tm_hour = ud.st.wHour;
		tmTemp.tm_mday = ud.st.wDay;
		tmTemp.tm_mon = ud.st.wMonth - 1;
		tmTemp.tm_year = ud.st.wYear - 1900;
		tmTemp.tm_wday = ud.st.wDayOfWeek;
		tmTemp.tm_yday = ud.wDayOfYear - 1;
		tmTemp.tm_isdst = 0;
		alloc(128);
		strftime(m_pData, 128, format, &tmTemp);
		return m_pData;
	}

	/*****************************************************************************************************************/

	inline wchar_t *fromDateTime(const wchar_t *format, const double &dt)
	{
		UDATE ud;
		if (S_OK != VarUdateFromDate(dt, 0, &ud))
			return 0;
		struct tm tmTemp;
		tmTemp.tm_sec = ud.st.wSecond;
		tmTemp.tm_min = ud.st.wMinute;
		tmTemp.tm_hour = ud.st.wHour;
		tmTemp.tm_mday = ud.st.wDay;
		tmTemp.tm_mon = ud.st.wMonth - 1;
		tmTemp.tm_year = ud.st.wYear - 1900;
		tmTemp.tm_wday = ud.st.wDayOfWeek;
		tmTemp.tm_yday = ud.wDayOfYear - 1;
		tmTemp.tm_isdst = 0;
		alloc(256);
		wcsftime((wchar_t *)m_pData, 128, format, &tmTemp);
		return (wchar_t *)m_pData;
	}

	/*****************************************************************************************************************/

#ifndef _UNICODE

	inline static double toDateTime(const char *dt)
	{
		double dtime;
		if (0 == dt)return -1;
		if (FAILED(VarDateFromStr((OLECHAR *)dt, MAKELANGID(LANG_VIETNAMESE, SUBLANG_DEFAULT), 0, &dtime)))
			return -1;
		return dtime;
	}

#else
	/*****************************************************************************************************************/

	inline static double toDateTime(const wchar_t *dt)
	{
		double dtime;
		if (0 == dt)return -1;
		if (FAILED(VarDateFromStr((OLECHAR *)dt, MAKELANGID(LANG_VIETNAMESE, SUBLANG_DEFAULT), 0, &dtime)))
			return -1;
		return dtime;
	}

#endif

	/************************************************************************/
	inline static CString UpcaseUnicode(CString szUnicode)
	{
		long length = szUnicode.GetLength();
		CString value;
		for (int i = 0; i < length; i++)
		{
			wchar_t  test = szUnicode.GetAt(i);
			CString test1 = (CString)test;

			if (test == 193 || test == 225)			//AS
				test1 = L"\xc1";
			else if (test == 192 || test == 224)	//AF
				test1 = L"\xc0";
			else if (test == 7842 || test == 7843)	//AR
				test1 = L"\x1ea2";
			else if (test == 195 || test == 227)	//AX
				test1 = L"\xc3";
			else if (test == 7840 || test == 7841)	//AJ
				test1 = L"\x1ea0";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 194 || test == 226)	//AA
				test1 = L"\xc2";
			else if (test == 7844 || test == 7845)	//AAS
				test1 = L"\x1ea4";
			else if (test == 7846 || test == 7847)	//AAF
				test1 = L"\x1ea6";
			else if (test == 7848 || test == 7849)	//AAR
				test1 = L"\x1ea8";
			else if (test == 7850 || test == 7851)	//AAX
				test1 = L"\x1eaa";
			else if (test == 7852 || test == 7853)	//AAJ
				test1 = L"\x1eac";
			else if (test == 258 || test == 259)	//AW
				test1 = L"\x102";
			else if (test == 7854 || test == 7855)	//AWS
				test1 = L"\x1eae";
			else if (test == 7856 || test == 7857)	//AWF
				test1 = L"\x1eb0";
			else if (test == 7858 || test == 7859)	//AWR
				test1 = L"\x1eb2";
			else if (test == 7860 || test == 7861)	//AWX
				test1 = L"\x1eb4";
			else if (test == 7862 || test == 7863)	//AWJ
				test1 = L"\x1eb6";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 201 || test == 233)	//ES
				test1 = L"\xc9";
			else if (test == 200 || test == 232)	//EF
				test1 = L"\xc8";
			else if (test == 7866 || test == 7867)	//ER
				test1 = L"\x1eba";
			else if (test == 7868 || test == 7869)	//EX
				test1 = L"\x1ebc";
			else if (test == 7864 || test == 7865)	//EJ
				test1 = L"\x1eb8";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 202 || test == 234)	//EE
				test1 = L"\xca";
			else if (test == 7870 || test == 7871)	//EES
				test1 = L"\x1ebe";
			else if (test == 7872 || test == 7873)	//EEF
				test1 = L"\x1ec0";
			else if (test == 7874 || test == 7875)	//EER
				test1 = L"\x1ec2";
			else if (test == 7876 || test == 7877)	//EEX
				test1 = L"\x1ec4";
			else if (test == 7878 || test == 7879)	//EEJ
				test1 = L"\x1ec6";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 211 || test == 243)	//OS
				test1 = L"\xd3";
			else if (test == 210 || test == 242)	//OF
				test1 = L"\xd2";
			else if (test == 7886 || test == 7887)	//OR
				test1 = L"\x1ece";
			else if (test == 213 || test == 245)	//OX
				test1 = L"\xd5";
			else if (test == 7884 || test == 7885)	//OJ
				test1 = L"\x1ecc";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 212 || test == 244)	//OO
				test1 = L"\xd4";
			else if (test == 7888 || test == 7889)	//OOS
				test1 = L"\x1ed0";
			else if (test == 7890 || test == 7891)	//OOF
				test1 = L"\x1ed2";
			else if (test == 7892 || test == 7893)	//OOR
				test1 = L"\x1ed4";
			else if (test == 7894 || test == 7895)	//OOX
				test1 = L"\x1ed6";
			else if (test == 7896 || test == 7897)	//OOJ
				test1 = L"\x1ed8";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 416 || test == 417)	//OW
				test1 = L"\x1a0";
			else if (test == 7898 || test == 7899)	//OWS
				test1 = L"\x1eda";
			else if (test == 7900 || test == 7901)	//OWF
				test1 = L"\x1edc";
			else if (test == 7902 || test == 7903)	//OWR
				test1 = L"\x1ede";
			else if (test == 7904 || test == 7905)	//OWX
				test1 = L"\x1ee0";
			else if (test == 7906 || test == 7907)	//OWJ 
				test1 = L"\x1ee2";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 218 || test == 250)	//US
				test1 = L"\xda";
			else if (test == 217 || test == 249)	//UF
				test1 = L"\xd9";
			else if (test == 7910 || test == 7911)	//UR
				test1 = L"\x1ee6";
			else if (test == 360 || test == 361)	//UX
				test1 = L"\x168";
			else if (test == 7908 || test == 7909)	//UJ 
				test1 = L"\x1ee4";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 431 || test == 432)	//UW
				test1 = L"\x1af";
			else if (test == 7912 || test == 7913)	//UWS
				test1 = L"\x1ee8";
			else if (test == 7914 || test == 7915)	//UWF
				test1 = L"\x1eea";
			else if (test == 7916 || test == 7917)	//UWR
				test1 = L"\x1eec";
			else if (test == 7918 || test == 7919)	//UWX
				test1 = L"\x1eee";
			else if (test == 7920 || test == 7921)	//UWJ
				test1 = L"\x1ef0";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 205 || test == 237)	//IS
				test1 = L"\xcd";
			else if (test == 204 || test == 236)	//IF
				test1 = L"\xcc";
			else if (test == 7880 || test == 7881)	//IR
				test1 = L"\x1ec8";
			else if (test == 296 || test == 297)		//IX
				test1 = L"\x128";
			else if (test == 7882 || test == 7883)	//IJ
				test1 = L"\x1eca";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 221 || test == 253)	//YS
				test1 = L"\xdd";
			else if (test == 7922 || test == 7923)	//YF
				test1 = L"\x1ef2";
			else if (test == 7926 || test == 7927)	//YR
				test1 = L"\x1ef6";
			else if (test == 7928 || test == 7929)	//YX
				test1 = L"\x1ef8";
			else if (test == 7924 || test == 7925)	//YJ
				test1 = L"\x1ef4";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 272 || test == 273)	//DD
				test1 = L"\x110";
			else test1.MakeUpper();
			value += test1;
		}
		return value;
	}
	inline static CString LowerUnicode(CString szUnicode)
	{
		long length = szUnicode.GetLength();
		CString value;
		for (int i = 0; i < length; i++)
		{
			wchar_t  test = szUnicode.GetAt(i);
			CString test1 = (CString)test;

			if (test == 193 || test == 225)			//AS
				test1 = L"\xE1";
			else if (test == 192 || test == 224)	//AF
				test1 = L"\xE0";
			else if (test == 7842 || test == 7843)	//AR
				test1 = L"\x1EA3";
			else if (test == 195 || test == 227)	//AX
				test1 = L"\xE3";
			else if (test == 7840 || test == 7841)	//AJ
				test1 = L"\x1EA1";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 194 || test == 226)	//AA
				test1 = L"\xE2";
			else if (test == 7844 || test == 7845)	//AAS
				test1 = L"\x1EA5";
			else if (test == 7846 || test == 7847)	//AAF
				test1 = L"\x1EA7";
			else if (test == 7848 || test == 7849)	//AAR
				test1 = L"\x1EA9";
			else if (test == 7850 || test == 7851)	//AAX
				test1 = L"\x1EAB";
			else if (test == 7852 || test == 7853)	//AAJ
				test1 = L"\x1EAD";
			else if (test == 258 || test == 259)	//AW
				test1 = L"\x103";
			else if (test == 7854 || test == 7855)	//AWS
				test1 = L"\x1EAF";
			else if (test == 7856 || test == 7857)	//AWF
				test1 = L"\x1EB1";
			else if (test == 7858 || test == 7859)	//AWR
				test1 = L"\x1EB3";
			else if (test == 7860 || test == 7861)	//AWX
				test1 = L"\x1EB5";
			else if (test == 7862 || test == 7863)	//AWJ
				test1 = L"\x1EB7";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 201 || test == 233)	//ES
				test1 = L"\xE9";
			else if (test == 200 || test == 232)	//EF
				test1 = L"\xE8";
			else if (test == 7866 || test == 7867)	//ER
				test1 = L"\x1EBB";
			else if (test == 7868 || test == 7869)	//EX
				test1 = L"\x1EBD";
			else if (test == 7864 || test == 7865)	//EJ
				test1 = L"\x1EA1";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 202 || test == 234)	//EE
				test1 = L"\xEA";
			else if (test == 7870 || test == 7871)	//EES
				test1 = L"\x1EBF";
			else if (test == 7872 || test == 7873)	//EEF
				test1 = L"\x1EC1";
			else if (test == 7874 || test == 7875)	//EER
				test1 = L"\x1EC3";
			else if (test == 7876 || test == 7877)	//EEX
				test1 = L"\x1EC5";
			else if (test == 7878 || test == 7879)	//EEJ
				test1 = L"\x1EC7";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 211 || test == 243)	//OS
				test1 = L"\xF3";
			else if (test == 210 || test == 242)	//OF
				test1 = L"\xF2";
			else if (test == 7886 || test == 7887)	//OR
				test1 = L"\x1ECF";
			else if (test == 213 || test == 245)	//OX
				test1 = L"\xF5";
			else if (test == 7884 || test == 7885)	//OJ
				test1 = L"\x1ECD";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 212 || test == 244)	//OO
				test1 = L"\xF4";
			else if (test == 7888 || test == 7889)	//OOS
				test1 = L"\x1ED1";
			else if (test == 7890 || test == 7891)	//OOF
				test1 = L"\x1ED3";
			else if (test == 7892 || test == 7893)	//OOR
				test1 = L"\x1ED5";
			else if (test == 7894 || test == 7895)	//OOX
				test1 = L"\x1ED7";
			else if (test == 7896 || test == 7897)	//OOJ
				test1 = L"\x1ED9";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 416 || test == 417)	//OW
				test1 = L"\x1A1";
			else if (test == 7898 || test == 7899)	//OWS
				test1 = L"\x1EDB";
			else if (test == 7900 || test == 7901)	//OWF
				test1 = L"\x1EDD";
			else if (test == 7902 || test == 7903)	//OWR
				test1 = L"\x1EDF";
			else if (test == 7904 || test == 7905)	//OWX
				test1 = L"\x1EE1";
			else if (test == 7906 || test == 7907)	//OWJ 
				test1 = L"\x1EE3";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 218 || test == 250)	//US
				test1 = L"\xFA";
			else if (test == 217 || test == 249)	//UF
				test1 = L"\xF9";
			else if (test == 7910 || test == 7911)	//UR
				test1 = L"\x1EE7";
			else if (test == 360 || test == 361)	//UX
				test1 = L"\x169";
			else if (test == 7908 || test == 7909)	//UJ 
				test1 = L"\x1EE5";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 431 || test == 432)	//UW
				test1 = L"\x1B0";
			else if (test == 7912 || test == 7913)	//UWS
				test1 = L"\x1EE9";
			else if (test == 7914 || test == 7915)	//UWF
				test1 = L"\x1EEB";
			else if (test == 7916 || test == 7917)	//UWR
				test1 = L"\x1EED";
			else if (test == 7918 || test == 7919)	//UWX
				test1 = L"\x1EEF";
			else if (test == 7920 || test == 7921)	//UWJ
				test1 = L"\x1EF1";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 205 || test == 237)	//IS
				test1 = L"\xED";
			else if (test == 204 || test == 236)	//IF
				test1 = L"\xEC";
			else if (test == 7880 || test == 7881)	//IR
				test1 = L"\x1EC9";
			else if (test == 296 || test == 297)		//IX
				test1 = L"\x129";
			else if (test == 7882 || test == 7883)	//IJ
				test1 = L"\x1ECB";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 221 || test == 253)	//YS
				test1 = L"\xFD";
			else if (test == 7922 || test == 7923)	//YF
				test1 = L"\x1EF3";
			else if (test == 7926 || test == 7927)	//YR
				test1 = L"\x1EF7";
			else if (test == 7928 || test == 7929)	//YX
				test1 = L"\x1EF9";
			else if (test == 7924 || test == 7925)	//YJ
				test1 = L"\x1EF5";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 272 || test == 273)	//DD
				test1 = L"\x111";
			else test1.MakeLower();
			value += test1;
		}
		return value;
	}
	/************************************************************************/
	inline static CString LowcaseUnicode(CString szUnicode)
	{
		long length = szUnicode.GetLength();
		CString value;
		for (int i = 0; i < length; i++)
		{
			wchar_t  test = szUnicode.GetAt(i);
			CString test1 = (CString)test;

			if (test == 193 || test == 225)			//AS
				test1 = L"\xe1";
			else if (test == 192 || test == 224)	//AF
				test1 = L"\xe0";
			else if (test == 7842 || test == 7843)	//AR
				test1 = L"\x1ea3";
			else if (test == 195 || test == 227)	//AX
				test1 = L"\xe3";
			else if (test == 7840 || test == 7841)	//AJ
				test1 = L"\x1ea1";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 194 || test == 226)	//AA
				test1 = L"\xe2";
			else if (test == 7844 || test == 7845)	//AAS
				test1 = L"\x1ea5";
			else if (test == 7846 || test == 7847)	//AAF
				test1 = L"\x1ea7";
			else if (test == 7848 || test == 7849)	//AAR
				test1 = L"\x1ea9";
			else if (test == 7850 || test == 7851)	//AAX
				test1 = L"\x1eab";
			else if (test == 7852 || test == 7853)	//AAJ
				test1 = L"\x1ead";
			else if (test == 258 || test == 259)	//AW
				test1 = L"\x103";
			else if (test == 7854 || test == 7855)	//AWS
				test1 = L"\x1eaf";
			else if (test == 7856 || test == 7857)	//AWF
				test1 = L"\x1eb1";
			else if (test == 7858 || test == 7859)	//AWR
				test1 = L"\x1eb3";
			else if (test == 7860 || test == 7861)	//AWX
				test1 = L"\x1eb5";
			else if (test == 7862 || test == 7863)	//AWJ
				test1 = L"\x1eb7";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 201 || test == 233)	//ES
				test1 = L"\xe9";
			else if (test == 200 || test == 232)	//EF
				test1 = L"\xe8";
			else if (test == 7866 || test == 7867)	//ER
				test1 = L"\x1ebb";
			else if (test == 7868 || test == 7869)	//EX
				test1 = L"\x1ebd";
			else if (test == 7864 || test == 7865)	//EJ
				test1 = L"\x1eb9";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 202 || test == 234)	//EE
				test1 = L"\xea";
			else if (test == 7870 || test == 7871)	//EES
				test1 = L"\x1ebf";
			else if (test == 7872 || test == 7873)	//EEF
				test1 = L"\x1ec1";
			else if (test == 7874 || test == 7875)	//EER
				test1 = L"\x1ec3";
			else if (test == 7876 || test == 7877)	//EEX
				test1 = L"\x1ec5";
			else if (test == 7878 || test == 7879)	//EEJ
				test1 = L"\x1ec7";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 211 || test == 243)	//OS
				test1 = L"\xf3	";
			else if (test == 210 || test == 242)	//OF
				test1 = L"\xf2";
			else if (test == 7886 || test == 7887)	//OR
				test1 = L"\x1ecf";
			else if (test == 213 || test == 245)	//OX
				test1 = L"\xf5";
			else if (test == 7884 || test == 7885)	//OJ
				test1 = L"\x1ecd";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 212 || test == 244)	//OO
				test1 = L"\xf4";
			else if (test == 7888 || test == 7889)	//OOS
				test1 = L"\x1ed1";
			else if (test == 7890 || test == 7891)	//OOF
				test1 = L"\x1ed3";
			else if (test == 7892 || test == 7893)	//OOR
				test1 = L"\x1ed5";
			else if (test == 7894 || test == 7895)	//OOX
				test1 = L"\x1ed7";
			else if (test == 7896 || test == 7897)	//OOJ
				test1 = L"\x1ed9";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 416 || test == 417)	//OW
				test1 = L"\x1a1";
			else if (test == 7898 || test == 7899)	//OWS
				test1 = L"\x1edb";
			else if (test == 7900 || test == 7901)	//OWF
				test1 = L"\x1edd";
			else if (test == 7902 || test == 7903)	//OWR
				test1 = L"\x1edf";
			else if (test == 7904 || test == 7905)	//OWX
				test1 = L"\x1ee1";
			else if (test == 7906 || test == 7907)	//OWJ 
				test1 = L"\x1ee3";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 218 || test == 250)	//US
				test1 = L"\xfa";
			else if (test == 217 || test == 249)	//UF
				test1 = L"\xf9";
			else if (test == 7910 || test == 7911)	//UR
				test1 = L"\x1ee7";
			else if (test == 360 || test == 361)	//UX
				test1 = L"\x169";
			else if (test == 7908 || test == 7909)	//UJ 
				test1 = L"\x1ee5";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 431 || test == 432)	//UW
				test1 = L"\x1b0	";
			else if (test == 7912 || test == 7913)	//UWS
				test1 = L"\x1ee9";
			else if (test == 7914 || test == 7915)	//UWF
				test1 = L"\x1eeb";
			else if (test == 7916 || test == 7917)	//UWR
				test1 = L"\x1eed";
			else if (test == 7918 || test == 7919)	//UWX
				test1 = L"\x1eef";
			else if (test == 7920 || test == 7921)	//UWJ
				test1 = L"\x1ef1";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 205 || test == 237)	//IS
				test1 = L"\xed";
			else if (test == 204 || test == 236)	//IF
				test1 = L"\xec";
			else if (test == 7880 || test == 7881)	//IR
				test1 = L"\x1ec9";
			else if (test == 296 || test == 297)		//IX
				test1 = L"\x129";
			else if (test == 7882 || test == 7883)	//IJ
				test1 = L"\x1ecb";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 221 || test == 253)	//YS
				test1 = L"\xfd";
			else if (test == 7922 || test == 7923)	//YF
				test1 = L"\x1ef3";
			else if (test == 7926 || test == 7927)	//YR
				test1 = L"\x1ef7";
			else if (test == 7928 || test == 7929)	//YX
				test1 = L"\x1ef9";
			else if (test == 7924 || test == 7925)	//YJ
				test1 = L"\x1ef5";
			//////////////////////////////////////////////////////////////////////////
			else if (test == 272 || test == 273)	//DD
				test1 = L"\x111";
			else test1.MakeLower();
			value += test1;
		}
		return value;

	}
};

#endif // !defined(AFX_STRUTIL_H__5F96D586_85C7_494E_90A9_D54201F65F7C__INCLUDED_)

