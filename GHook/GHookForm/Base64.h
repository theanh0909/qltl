#pragma once

string base64_encode(unsigned char const*, unsigned int len);
string base64_decode(string const& s);
CString BaseEncodeEx(CString sValue, int itype = 0);
CString BaseDecodeEx(CString sValue);

