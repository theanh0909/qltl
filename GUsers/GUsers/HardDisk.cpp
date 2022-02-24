#include "pch.h"
#include "HardDisk.h"
#include "diskid32.h"

#define dfENCRYPT_NUM	88
#define dfSEG_LEN		5

BOOL continue_work = TRUE;
int stasDisk = 0;

HardDisk::HardDisk()
{
}

HardDisk::~HardDisk()
{
}

CString HardDisk::_CheckHardWareID()
{
	try
	{
		char lpszval[MAX_LEN];
		char localKeys[MAX_LEN];
		sprintf(lpszval, "");

		if (!ReadPhysicalDriveInNTWithZeroRights(0, lpszval))
			if (!ReadPhysicalDriveInNTUsingSmart(0, lpszval))
				if (!ReadIdeDriveAsScsiDriveInNT(0, lpszval))
					if (!ReadPhysicalDriveInNTWithAdminRights(0, lpszval))
					{
						continue_work = FALSE;
					}

		PadLeft(lpszval, 32, ' ');
		if (!EncryptSig(lpszval, localKeys, dfENCRYPT_NUM))
		{
			continue_work = FALSE;
			return L"";
		}

		return _ConvertCharsToCstring(localKeys);
	}
	catch (...) {}
	return L"";
}

CString HardDisk::_ConvertCharsToCstring(char* chr)
{
	int num = MultiByteToWideChar(CP_UTF8, 0, chr, -1, NULL, 0);
	wchar_t* wstr = new wchar_t[num + 1];
	MultiByteToWideChar(CP_UTF8, 0, chr, -1, wstr, num);
	wstr[num] = '\0';
	CString szval = L"";
	szval.Format(L"%s", wstr);
	return szval;
}

bool HardDisk::EncryptSig(char *strpass, char* strouput, char num)
{
	char temp[MAX_LEN];
	char strinput[MAX_LEN];
	char encryptdata[MAX_LEN];
	
	memset(encryptdata, 0, MAX_LEN);
	memset(temp, 0, MAX_LEN);

	if (!strpass || !strouput) return false;

	RemoveDashes(strpass);
	strcpy(strinput, strpass);
	int length = strlen(strinput);
	InvertSig(strinput, length);

	ScrambleString(strinput, encryptdata, num, length);

	InvertSig(encryptdata, length);
	ConvertToHexs(encryptdata, temp, length);
	Pick(temp, 5 * dfSEG_LEN);
	AddDashes(temp, strouput);
	return true;
}

void HardDisk::Pick(char *str, int len) 
{
	if ((int)strlen(str) < len * 2) return;
	for (int i = 0; i < len; i++) str[i] = str[i * 2];
	str[len] = 0;
}

bool HardDisk::AddDashes(char *strInput, char *strOutput) 
{
	if (!strInput || !strOutput) return false;

	int len = strlen(strInput);
	int offset = len % dfSEG_LEN;
	int i = 0, j = 0;
	while (strInput[i] != NULL) 
	{
		if ((i - offset) % dfSEG_LEN == 0 && i > 0) strOutput[j++] = '-';
		strOutput[j++] = strInput[i++];
	}
	strOutput[j] = NULL;
	return true;
}

bool HardDisk::ConvertToHexs(char* strChars, char* strHex, int length)
{
	char *strDest;
	if (!strChars || !strChars) return false;

	memset(strHex, 0, MAX_LEN);
	strDest = strHex;
	for (int i = 0; i < length; i++) 
	{
		unsigned char ch = strChars[i];
		sprintf(strDest, "%02X", ch);
		strDest += 2;
	}

	return true;
}

bool HardDisk::ScrambleString(char * strinput, char * encryptdata, char num, int length)
{
	for (int counter = 0; counter < length; counter++) 
	{
		int tmp = strinput[counter];
		if (counter > 0) tmp = tmp + encryptdata[counter - 1];
		else 
		{
			if (counter < length - 1) tmp = tmp + strinput[counter + 1];
		}

		if (counter % 2) tmp += num;
		else tmp -= num;

		tmp = tmp ^ (10 - num);
		encryptdata[counter] = tmp;
	}
	return true;
}

bool HardDisk::InvertSig(char *strpass, int length)
{
	char tmp;
	int mid = (length - 1) / 2;
	for (int i = 0; i <= mid; i++)
	{
		tmp = strpass[i];
		strpass[i] = strpass[length - i - 1];
		strpass[length - i - 1] = tmp;
	}
	return true;
}

bool HardDisk::RemoveDashes(char *strInput) 
{
	if (!strInput) return false;

	if ((int)(strchr(strInput, '-') - strInput + 1) != NULL) 
	{
		int i = 0, j = 0;
		while (strInput[i] != NULL) 
		{
			if (strInput[i] != '-') strInput[j++] = strInput[i];
			i++;
		}
		strInput[j] = NULL;
	}

	return true;
}

void HardDisk::PadLeft(char *str, int length, char c) 
{
	if ((int)strlen(str) >= length) return;

	int diff = length - strlen(str);
	for (int i = strlen(str); i >= 0; i--) str[i + diff] = str[i];
	for (int i = 0; i < diff; i++) str[i] = c;		
}
