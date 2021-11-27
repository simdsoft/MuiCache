#include <Windows.h>

static BOOL GetNtVersionNumbers(LPOSVERSIONINFO lpVersionInformation)
{
  BOOL bRet = FALSE;
  if (lpVersionInformation == NULL)
    return bRet;

  HMODULE hModNtdll = ::GetModuleHandle(L"ntdll.dll");
  if (hModNtdll)
  {
    typedef void(WINAPI * pfRTLGETNTVERSIONNUMBERS)(DWORD*, DWORD*, DWORD*);
    pfRTLGETNTVERSIONNUMBERS pfRtlGetNtVersionNumbers;
    pfRtlGetNtVersionNumbers = (pfRTLGETNTVERSIONNUMBERS)::GetProcAddress(hModNtdll, "RtlGetNtVersionNumbers");
    if (pfRtlGetNtVersionNumbers)
    {
      pfRtlGetNtVersionNumbers(&lpVersionInformation->dwMajorVersion, &lpVersionInformation->dwMinorVersion,
                               &lpVersionInformation->dwBuildNumber);
      lpVersionInformation->dwBuildNumber &= 0x00ffffff;
      bRet = TRUE;
    }
  }

  return bRet;
}

extern "C" BOOL IsWindows10OrGreater() {
	OSVERSIONINFO info;
	return GetNtVersionNumbers(&info) && info.dwBuildNumber >= 10240;
}
