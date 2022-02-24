#pragma once

extern BOOL continue_work;
extern int stasDisk;

class HardDisk
{
public:
	HardDisk();
	~HardDisk();

	CString _CheckHardWareID();
	CString _ConvertCharsToCstring(char* chr);
	bool EncryptSig(char *strinput, char* strouput, char num);
	void Pick(char *str, int len);
	bool AddDashes(char *strInput, char *strOutput);
	bool ConvertToHexs(char* strChars, char* strHex, int length);
	bool ScrambleString(char * strinput, char * encryptdata, char num, int length);
	bool InvertSig(char *strpass, int length);
	bool RemoveDashes(char *strInput);
	void PadLeft(char *str, int length, char c);
};

