// nsFileVSI.cpp : Defines the exported functions for the DLL application.
// V1.3: Add TaskPin support
// V1.2: First release
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <winreg.h>
#include <shlwapi.h>
#include <ShellAPI.h>
#include "nsis/pluginapi.h" // nsis plugin

#if defined(_DEBUG)
#define MUICACHE_WAIT_DEBUGGER 1
#pragma comment(lib, "Shlwapi.lib")
#endif

#define MAX_KEY_LENGTH 255
#define MAX_VALUE_NAME 16383

#define TB_PIN_OK   42
#define TB_PIN_FAIL 31

extern HRESULT TaskbarSetPinState(LPCTSTR pszPath, BOOL pinning);
extern BOOL IsWindows10OrGreater();
extern BOOL SetShortcutAppId(LPCTSTR shortcut, LPCTSTR appid);

static void MuiCache_Clear(const wchar_t* imageName, HKEY hRegRoot, const wchar_t* cacheRegPath) {

    HKEY hKeyMuiCache;
    //TCHAR    achKey[MAX_KEY_LENGTH];   // buffer for subkey name
    //DWORD    cbName;                   // size of name string 
    TCHAR    achClass[MAX_PATH] = TEXT("");  // buffer for class name 
    DWORD    cchClassName = MAX_PATH;  // size of class string 
    DWORD    cSubKeys = 0;               // number of subkeys 
    DWORD    cbMaxSubKey;              // longest subkey size 
    DWORD    cchMaxClass;              // longest class string 
    DWORD    cValues;              // number of values for key 
    DWORD    cchMaxValue;          // longest value name 
    DWORD    cbMaxValueData;       // longest value data 
    DWORD    cbSecurityDescriptor; // size of security descriptor 
    FILETIME ftLastWriteTime;      // last write time 

    int i, n;
    LSTATUS retCode;

    static TCHAR  achValue[MAX_VALUE_NAME];
    DWORD cchValue = MAX_VALUE_NAME;

#if defined(_DEBUG)
    MessageBoxW(NULL, L"Waiting for debugger to attach...", L"Waiting for debugger to attach...", MB_OK | MB_ICONEXCLAMATION);
#endif

    if (RegOpenKeyEx(hRegRoot, cacheRegPath, 0, KEY_READ | KEY_WRITE, &hKeyMuiCache) == 0)
    {
        retCode = RegQueryInfoKey(
            hKeyMuiCache,            // key handle 
            achClass,                // buffer for class name 
            &cchClassName,           // size of class string 
            NULL,                    // reserved 
            &cSubKeys,               // number of subkeys 
            &cbMaxSubKey,            // longest subkey size 
            &cchMaxClass,            // longest class string 
            &cValues,                // number of values for this key 
            &cchMaxValue,            // longest value name 
            &cbMaxValueData,         // longest value data 
            &cbSecurityDescriptor,   // security descriptor 
            &ftLastWriteTime);       // last write time 

        // Enumerate the subkeys, until RegEnumKeyEx fails.

        // Enumerate the key values. 
        if (cValues)
        {
            n = (int)cValues;
            for (i = 0, retCode = ERROR_SUCCESS; i < n; )
            {
                cchValue = MAX_VALUE_NAME;
                achValue[0] = '\0';
                retCode = RegEnumValue(hKeyMuiCache, i,
                    achValue,
                    &cchValue,
                    NULL,
                    NULL,
                    NULL,
                    NULL);

                if (retCode == ERROR_SUCCESS && StrStrW(achValue, imageName))
                {
                    RegDeleteValue(hKeyMuiCache, achValue);
                    --n;
                    continue;
                }

                ++i;
            }
        }

        RegCloseKey(hKeyMuiCache);
    }
}

#if defined(__cplusplus)
extern "C" {
#endif
    // To work with Unicode version of NSIS, please use TCHAR-type
    // functions for accessing the variables and the stack.
    void __declspec(dllexport) Clear(HWND hwndParent, int string_size,
        LPTSTR variables, stack_t** stacktop,
        extra_parameters* extra, ...)
    {
        // note if you want parameters from the stack, pop them off in order.
        // i.e. if you are called via exdll::myFunction file.dat read.txt
        // calling popstring() the first time would give you file.dat,
        // and the second time would give you read.txt. 
        // you should empty the stack of your parameters, and ONLY your
        // parameters.
        TCHAR appImageName[256];
        EXDLL_INIT();

        popstring(appImageName);
        MuiCache_Clear(appImageName, HKEY_CLASSES_ROOT, L"Local Settings\\Software\\Microsoft\\Windows\\Shell\\MuiCache");
		// MuiCache_Clear(appImageName, HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\UFH\\SHC");
		// HKEY_USERS\S-1-5-21-3324583540-2673638656-2559637184-1002\Software\Microsoft\Windows\CurrentVersion\UFH\SHC
		
		// MuiCache_Clear(appImageName, HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\Services\\SharedAccess\\Parameters\\FirewallPolicy\\FirewallRules");
    }

	void __declspec(dllexport) TaskbarUnpin(HWND hwndParent, int string_size,
        LPTSTR variables, stack_t** stacktop,
        extra_parameters* extra, ...)
    {
        // note if you want parameters from the stack, pop them off in order.
        // i.e. if you are called via exdll::myFunction file.dat read.txt
        // calling popstring() the first time would give you file.dat,
        // and the second time would give you read.txt. 
        // you should empty the stack of your parameters, and ONLY your
        // parameters.
        TCHAR shortcut[512];
		INT_PTR result;
        EXDLL_INIT();

        popstring(shortcut);
		result = (INT_PTR)ShellExecute(NULL, L"taskbarunpin", shortcut, NULL, NULL, 0);
		pushint(result);
    }

	void __declspec(dllexport) TaskbarPin(HWND hwndParent, int string_size,
		LPTSTR variables, stack_t** stacktop,
		extra_parameters* extra, ...)
    {
        // note if you want parameters from the stack, pop them off in order.
        // i.e. if you are called via exdll::myFunction file.dat read.txt
        // calling popstring() the first time would give you file.dat,
        // and the second time would give you read.txt. 
        // you should empty the stack of your parameters, and ONLY your
        // parameters.
        TCHAR shortcut[512];
		INT_PTR result;
        EXDLL_INIT();

        popstring(shortcut);
		if(IsWindows10OrGreater())
			result = TaskbarSetPinState(shortcut, TRUE) == S_OK ? TB_PIN_OK : TB_PIN_FAIL;
		else
			result = (INT_PTR)ShellExecute(NULL, L"taskbarpin", shortcut, NULL, NULL, 0);
		pushint(result);
    }

	void __declspec(dllexport) SetLnkAppId(HWND hwndParent, int string_size,
		LPTSTR variables, stack_t** stacktop,
		extra_parameters* extra, ...)
    {
        // note if you want parameters from the stack, pop them off in order.
        // i.e. if you are called via exdll::myFunction file.dat read.txt
        // calling popstring() the first time would give you file.dat,
        // and the second time would give you read.txt. 
        // you should empty the stack of your parameters, and ONLY your
        // parameters.
        TCHAR shortcut[512];
		TCHAR lnk[512];
		// INT_PTR result;
        EXDLL_INIT();

        popstring(shortcut);
		popstring(lnk);

		SetShortcutAppId(shortcut, lnk);
		// pushint(result);
    }


    BOOL WINAPI DllMain(HINSTANCE hInst, ULONG ul_reason_for_call, LPVOID lpReserved)
    {
        // Perform actions based on the reason for calling.
        switch (ul_reason_for_call)
        {
        case DLL_PROCESS_ATTACH:
            // Initialize once for each new process.
            // Return FALSE to fail DLL load.
#if defined(MUICACHE_WAIT_DEBUGGER)
            MessageBox(NULL, L"Waiting for visual studio debugger to attach...", L"Waiting for visual studio debugger to attach...", MB_OK|MB_ICONEXCLAMATION);
#endif
            break;

        case DLL_THREAD_ATTACH:
            // Do thread-specific initialization.
            break;

        case DLL_THREAD_DETACH:
            // Do thread-specific cleanup.
            break;

        case DLL_PROCESS_DETACH:
            // Perform any necessary cleanup.
            break;
        }
        return TRUE;  // Successful DLL_PROCESS_ATTACH.
    }
#if defined(__cplusplus)
}
#endif
