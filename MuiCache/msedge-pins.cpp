/*
Copyright (c) 2020 by Gee Law

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

references:
  - https://geelaw.blog/entries/msedge-pins/
  - https://geelaw.blog/entries/msedge-pins/assets/pin.cc
**/
#if defined(_DEBUG)
#pragma comment(lib, "ole32")
#pragma comment(lib, "shell32")
#endif

#define STRICT_TYPED_ITEMIDS
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <objbase.h>
#include <shlobj.h>

const GUID CLSID_TaskbandPin = { 0x90aa3a4e, 0x1cba, 0x4233, {0xb8, 0xbb, 0x53, 0x57, 0x73, 0xd4, 0x84, 0x49} };
const GUID IID_IPinnedList3 = { 0x0dd79ae2, 0xd156, 0x45d4, {0x9e, 0xeb, 0x3b, 0x54, 0x97, 0x69, 0xe9, 0x40} };
enum PLMC
{
    PLMC_EXPLORER = 4
};

typedef struct IPinnedList3Vtbl IPinnedList3Vtbl;
struct IPinnedList3 {
    IPinnedList3Vtbl* vtbl;
};

typedef ULONG (STDMETHODCALLTYPE *ReleaseFuncPtr)(struct IPinnedList3* that);
typedef HRESULT (STDMETHODCALLTYPE *ModifyFuncPtr)(struct IPinnedList3* that, PCIDLIST_ABSOLUTE unpin, PCIDLIST_ABSOLUTE pin, int);

struct IPinnedList3Vtbl {
    void* QueryInterface;
    void* AddRef;
    ReleaseFuncPtr Release;
    void* MethodSlot4;
    void* MethodSlot5;
    void* MethodSlot6;
    void* MethodSlot7;
    void* MethodSlot8;
    void* MethodSlot9;
    void* MethodSlot10;
    void* MethodSlot11;
    void* MethodSlot12;
    void* MethodSlot13;
    void* MethodSlot14;
    void* MethodSlot15;
    void* MethodSlot16;
    ModifyFuncPtr Modify;
};
extern "C" HRESULT TaskbarSetPinState(LPCTSTR pszPath, BOOL pinning)
{
    HRESULT hr;
    PIDLIST_ABSOLUTE pidl;
	struct IPinnedList3* pinnedList;

    hr = CoInitialize(NULL);
    if (!SUCCEEDED(hr))
        return hr;

    do {
        pidl = NULL;
        pidl = ILCreateFromPath(pszPath);
        if (!pidl)
            break;

        pinnedList = NULL;
        hr = CoCreateInstance(CLSID_TaskbandPin, NULL, CLSCTX_ALL, IID_IPinnedList3, (LPVOID*)(&pinnedList));
        if (!SUCCEEDED(hr))
            break;

		hr = pinnedList->vtbl->Modify(pinnedList, pinning ? NULL : pidl, pinning ? pidl : NULL, PLMC_EXPLORER);
        pinnedList->vtbl->Release(pinnedList);
    } while (0);

    if (pidl)
        ILFree(pidl);

    CoUninitialize();
    return hr;
}
