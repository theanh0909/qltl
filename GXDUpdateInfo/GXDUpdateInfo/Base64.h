#pragma once
class Base64Ex
{
public:
	Base64Ex();
	~Base64Ex();

	int RandomTime();
	CString ConvertstringtoUTF8(string ret);
	string ConvertUTF8tostring(CString szval);	

	string base64_encode(unsigned char const*, unsigned int len);
	string base64_decode(string const& s);

	CString BaseEncodeEx(CString sValue, int itype = 0);
	CString BaseDecodeEx(CString sValue);
};

