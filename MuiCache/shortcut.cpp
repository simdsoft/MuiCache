// Copyright (c) 2012 The Chromium Authors. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "shortcut.h"

#if defined(_DEBUG)
#pragma comment(lib, "Propsys.lib")
#endif

extern "C" void* __cdecl memset(void *p, int c, size_t z);

static bool is_file_exists(LPCTSTR filename)
{
  WIN32_FILE_ATTRIBUTE_DATA attrs = {0};
  return GetFileAttributesEx(filename, GetFileExInfoStandard, &attrs) &&
         !(attrs.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
}

/// CLASS ComPtr
#define __WRL_CLASSIC_COM__

namespace Details
{
struct BoolStruct
{
    int Member;
};

// Helper object that makes IUnknown methods private
template <typename T>
class RemoveIUnknownBase : public T
{
private:
    ~RemoveIUnknownBase();

    // STDMETHOD macro implies virtual.
    // ComPtr can be used with any class that implements the 3 methods of IUnknown.
    // ComPtr does not require these methods to be virtual.
    // When ComPtr is used with a class without a virtual table, marking the functions
    // as virtual in this class adds unnecessary overhead.

    // WRL::ComPtr is designed to take over responsibility for managing
    // COM ref-counts, reducing the risk of leaking references to COM objects
    // (implying memory leaks) or of over-releasing (leading to crashes and/or
    // heap corruption).  To maximize the safey of ComPtr-using code, RemoveIUnknown
    // is used to hide from ComPtr clients direct access to the native IUnknown
    // refcounting methods.  However, RemoveIUnknown is only used in checked/debug builds.
    // Its implementation is incompatible with the performance win that
    // comes from marking WRL-implemented runtimeclasses as "final" (or "sealed").
    // That performance win is enabled in production/release builds, while
    // ComPtr's use of RemoveIUnknown is enabled in check/debug builds.
    // This difference in compile-time behavior between production/release and
    // checked/debug builds leads to developer friction for any projects
    // that do production builds more frequently than they do checked/debug builds;
    // code that compiles in production builds fails in checked/debug builds.
    // Experience has shown that ComPtr clients attempt to access
    // IUnknown::QueryInterface far more commonly than they do IUnknown::AddRef
    // or IUnknown::Release.  And that same experience has shown that they tend
    // to use QueryInterface in a safe way, appropriately using a ComPtr to
    // provide the output-paremeter to QueryInteface.  Based on that experience,
    // RemoveIUnknown will no longer block access to QueryInterface.  (It is worth
    // noting, however, that there is never an actual requirement to access
    // QueryInterface directly; any caller could instead use a call
    // to CompPtr::CopyTo(REFIID riid, void** ptr).)
    //
    // HRESULT __stdcall QueryInterface(REFIID riid, _COM_Outptr_ void **ppvObject);
    ULONG __stdcall AddRef();
    ULONG __stdcall Release();
};

template<typename T>
struct RemoveIUnknown
{
    typedef RemoveIUnknownBase<T> ReturnType;
};

template<typename T>
struct RemoveIUnknown<const T>
{
    typedef const RemoveIUnknownBase<T> ReturnType;
};

template <typename T> // T should be the ComPtr<T> or a derived type of it, not just the interface
class ComPtrRefBase
{
public:
    typedef typename T::InterfaceType InterfaceType;

#ifndef __WRL_CLASSIC_COM__
    operator IInspectable**() const
    {
        //static_assert(__is_base_of(IInspectable, InterfaceType), "Invalid cast: InterfaceType does not derive from IInspectable");
        return reinterpret_cast<IInspectable**>(ptr_->ReleaseAndGetAddressOf());
    }
#endif

    operator IUnknown**() const
    {
        //static_assert(__is_base_of(IUnknown, InterfaceType), "Invalid cast: InterfaceType does not derive from IUnknown");
        return reinterpret_cast<IUnknown**>(ptr_->ReleaseAndGetAddressOf());
    }

protected:
    T* ptr_;
};

template <typename T>
class ComPtrRef : public Details::ComPtrRefBase<T> // T should be the ComPtr<T> or a derived type of it, not just the interface
{
    typedef Details::ComPtrRefBase<T> Super;
    typedef typename Super::InterfaceType InterfaceType;
public:
    ComPtrRef(_In_opt_ T* ptr)
    {
         this->ptr_ = ptr;
    }

    // Conversion operators
    operator void**() const
    {
        return reinterpret_cast<void**>(this->ptr_->ReleaseAndGetAddressOf());
    }

    // This is our operator ComPtr<U> (or the latest derived class from ComPtr (e.g. WeakRef))
    operator T*()
    {
        *this->ptr_ = nullptr;
        return this->ptr_;
    }

    // We define operator InterfaceType**() here instead of on ComPtrRefBase<T>, since
    // if InterfaceType is IUnknown or IInspectable, having it on the base will collide.
    operator InterfaceType**()
    {
        return this->ptr_->ReleaseAndGetAddressOf();
    }

    // This is used for IID_PPV_ARGS in order to do __uuidof(**(ppType)).
    // It does not need to clear  ptr_ at this point, it is done at IID_PPV_ARGS_Helper(ComPtrRef&) later in this file.
    InterfaceType* operator *()
    {
        return this->ptr_->Get();
    }

    // Explicit functions
    InterfaceType* const * GetAddressOf() const
    {
        return this->ptr_->GetAddressOf();
    }

    InterfaceType** ReleaseAndGetAddressOf()
    {
        return this->ptr_->ReleaseAndGetAddressOf();
    }
};

} // namespace Details

template <typename T>
class ComPtr
{
public:
    typedef T InterfaceType;

protected:
    InterfaceType *ptr_;
    template<class U> friend class ComPtr;

    void InternalAddRef() const
    {
        if (ptr_ != nullptr)
        {
            ptr_->AddRef();
        }
    }

    unsigned long InternalRelease()
    {
        unsigned long ref = 0;
        T* temp = ptr_;

        if (temp != nullptr)
        {
            ptr_ = nullptr;
            ref = temp->Release();
        }

        return ref;
    }

public:
#pragma region constructors
    ComPtr() : ptr_(nullptr)
    {
    }

    ComPtr(decltype(__nullptr)) : ptr_(nullptr)
    {
    }

    template<class U>
    ComPtr(_In_opt_ U *other) : ptr_(other)
    {
        InternalAddRef();
    }

    ComPtr(const ComPtr& other) : ptr_(other.ptr_)
    {
        InternalAddRef();
    }

    // copy constructor that allows to instantiate class when U* is convertible to T*
    template<class U>
    ComPtr(const ComPtr<U> &other) :
        ptr_(other.ptr_)
    {
        InternalAddRef();
    }

    //ComPtr(_Inout_ ComPtr &&other) : ptr_(nullptr)
    //{
    //    if (this != reinterpret_cast<ComPtr*>(&reinterpret_cast<unsigned char&>(other)))
    //    {
    //        Swap(other);
    //    }
    //}

    // Move constructor that allows instantiation of a class when U* is convertible to T*
    //template<class U>
    //ComPtr(_Inout_ ComPtr<U>&& other, typename Details::EnableIf<Details::IsConvertible<U*, T*>::value, void *>::type * = 0) :
    //    ptr_(other.ptr_)
    //{
    //    other.ptr_ = nullptr;
    //}
#pragma endregion

#pragma region destructor
    ~ComPtr()
    {
        InternalRelease();
    }
#pragma endregion

#pragma region assignment
    ComPtr& operator=(decltype(__nullptr))
    {
        InternalRelease();
        return *this;
    }

    ComPtr& operator=(_In_opt_ T *other)
    {
        if (ptr_ != other)
        {
            ComPtr(other).Swap(*this);
        }
        return *this;
    }

    template <typename U>
    ComPtr& operator=(_In_opt_ U *other)
    {
        ComPtr(other).Swap(*this);
        return *this;
    }

    ComPtr& operator=(const ComPtr &other)
    {
        if (ptr_ != other.ptr_)
        {
            ComPtr(other).Swap(*this);
        }
        return *this;
    }

    template<class U>
    ComPtr& operator=(const ComPtr<U>& other)
    {
        ComPtr(other).Swap(*this);
        return *this;
    }

    //ComPtr& operator=(_Inout_ ComPtr &&other)
    //{
    //    ComPtr(static_cast<ComPtr&&>(other)).Swap(*this);
    //    return *this;
    //}

    //template<class U>
    //ComPtr& operator=(_Inout_ ComPtr<U>&& other)
    //{
    //    ComPtr(static_cast<ComPtr<U>&&>(other)).Swap(*this);
    //    return *this;
    //}
#pragma endregion

#pragma region modifiers
    //void Swap(_Inout_ ComPtr&& r)
    //{
    //    T* tmp = ptr_;
    //    ptr_ = r.ptr_;
    //    r.ptr_ = tmp;
    //}

    void Swap(_Inout_ ComPtr& r)
    {
        T* tmp = ptr_;
        ptr_ = r.ptr_;
        r.ptr_ = tmp;
    }
#pragma endregion

    operator bool() const
    {
        return Get() != nullptr;
    }

    T* Get() const
    {
        return ptr_;
    }
    

    InterfaceType* operator->() const
    {
        return ptr_;
    }

    Details::ComPtrRef<ComPtr<T>> operator&()
    {
        return Details::ComPtrRef<ComPtr<T>>(this);
    }

    const Details::ComPtrRef<const ComPtr<T>> operator&() const
    {
        return Details::ComPtrRef<const ComPtr<T>>(this);
    }

    T* const* GetAddressOf() const
    {
        return &ptr_;
    }

    T** GetAddressOf()
    {
        return &ptr_;
    }

    T** ReleaseAndGetAddressOf()
    {
        InternalRelease();
        return &ptr_;
    }

    T* Detach()
    {
        T* ptr = ptr_;
        ptr_ = nullptr;
        return ptr;
    }

    void Attach(_In_opt_ InterfaceType* other)
    {
        if (ptr_ != nullptr)
        {
            auto ref = ptr_->Release();
            DBG_UNREFERENCED_LOCAL_VARIABLE(ref);
            // Attaching to the same object only works if duplicate references are being coalesced. Otherwise
            // re-attaching will cause the pointer to be released and may cause a crash on a subsequent dereference.
            //__WRL_ASSERT__(ref != 0 || ptr_ != other);
        }

        ptr_ = other;
    }

    unsigned long Reset()
    {
        return InternalRelease();
    }

    // query for U interface
    template<typename U>
    HRESULT As(_Inout_ Details::ComPtrRef<ComPtr<U>> p) const
    {
        return ptr_->QueryInterface(__uuidof(U), p);
    }

    // query for U interface
    template<typename U>
    HRESULT As(_Out_ ComPtr<U>* p) const
    {
        return ptr_->QueryInterface(__uuidof(U), reinterpret_cast<void**>(p->ReleaseAndGetAddressOf()));
    }

    // query for riid interface and return as IUnknown
    HRESULT AsIID(REFIID riid, _Out_ ComPtr<IUnknown>* p) const
    {
        return ptr_->QueryInterface(riid, reinterpret_cast<void**>(p->ReleaseAndGetAddressOf()));
    }
};    // ComPtr


class ScopedPropVariant {
public:
  ScopedPropVariant() { PropVariantInit(&pv_); }
  ScopedPropVariant(const ScopedPropVariant&);
  ScopedPropVariant& operator=(const ScopedPropVariant&) ;
  ~ScopedPropVariant() { Reset(); }
  // Returns a pointer to the underlying PROPVARIANT for use as an out param in
  // a function call.
  PROPVARIANT* Receive()
  {
    //DCHECK_EQ(pv_.vt, VT_EMPTY);
    return &pv_;
  }
  // Clears the instance to prepare it for re-use (e.g., via Receive).
  void Reset()
  {
    if (pv_.vt != VT_EMPTY)
    {
      HRESULT result = PropVariantClear(&pv_);
      //DCHECK_EQ(result, S_OK);
    }
  }
  const PROPVARIANT& get() const { return pv_; }
  const PROPVARIANT* ptr() const { return &pv_; }

private:
  PROPVARIANT pv_;
  // Comparison operators for ScopedPropVariant are not supported at this point.
  bool operator==(const ScopedPropVariant&) const;
  bool operator!=(const ScopedPropVariant&) const;
};


namespace
{
// Sets the value of |property_key| to |property_value| in |property_store|.
bool SetPropVariantValueForPropertyStore(IPropertyStore* property_store, const PROPERTYKEY& property_key, const ScopedPropVariant& property_value)
{
  // DCHECK(property_store);
  HRESULT result = property_store->SetValue(property_key, property_value.get());
  if (result == S_OK)
    result = property_store->Commit();
  if (SUCCEEDED(result))
    return true;
  // #if DCHECK_IS_ON()
  if (HRESULT_FACILITY(result) == FACILITY_WIN32)
    ::SetLastError(HRESULT_CODE(result));
  // See third_party/perl/c/i686-w64-mingw32/include/propkey.h for GUID and
  // PID definitions.
  // DPLOG(ERROR) << "Failed to set property with GUID " << WStringFromGUID(property_key.fmtid) << " PID " << property_key.pid;
  // #endif
  return false;
}

bool SetStringValueForPropertyStore(IPropertyStore* property_store, const PROPERTYKEY& property_key, const wchar_t* property_string_value)
{
  ScopedPropVariant property_value;
  if (FAILED(InitPropVariantFromString(property_string_value, property_value.Receive())))
  {
    return false;
  }
  return SetPropVariantValueForPropertyStore(property_store, property_key, property_value);
}

bool SetAppIdForPropertyStore(IPropertyStore* property_store, const wchar_t* app_id)
{
  // App id should be less than 128 chars and contain no space. And recommended
  // format is CompanyName.ProductName[.SubProduct.ProductNumber].
  // See
  // https://docs.microsoft.com/en-us/windows/win32/shell/appids#how-to-form-an-application-defined-appusermodelid
  // DCHECK_LT(lstrlen(app_id), 128);
  // DCHECK_EQ(wcschr(app_id, L' '), nullptr);
  return SetStringValueForPropertyStore(property_store, PKEY_AppUserModel_ID, app_id);
}
bool SetBooleanValueForPropertyStore(IPropertyStore* property_store, const PROPERTYKEY& property_key, bool property_bool_value)
{
  ScopedPropVariant property_value;
  if (FAILED(InitPropVariantFromBoolean(property_bool_value, property_value.Receive())))
  {
    return false;
  }
  return SetPropVariantValueForPropertyStore(property_store, property_key, property_value);
}
bool SetClsidForPropertyStore(IPropertyStore* property_store, const PROPERTYKEY& property_key, const CLSID& property_clsid_value)
{
  ScopedPropVariant property_value;
  if (FAILED(InitPropVariantFromCLSID(property_clsid_value, property_value.Receive())))
  {
    return false;
  }
  return SetPropVariantValueForPropertyStore(property_store, property_key, property_value);
}
} // namespace

namespace base
{
namespace win
{

namespace
{

// Initializes |i_shell_link| and |i_persist_file| (releasing them first if they
// are already initialized).
// If |shortcut| is not NULL, loads |shortcut| into |i_persist_file|.
// If any of the above steps fail, both |i_shell_link| and |i_persist_file| will
// be released.
void InitializeShortcutInterfaces(LPCTSTR shortcut, ComPtr<IShellLink>* i_shell_link, ComPtr<IPersistFile>* i_persist_file)
{
  // Reset in the inverse order of acquisition.
  i_persist_file->Reset();
  i_shell_link->Reset();

  
  ComPtr<IShellLink> shell_link;
  if (FAILED(::CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(static_cast<IShellLink**>(&shell_link)))))
  {
    return;
  }
  ComPtr<IPersistFile> persist_file;
  if (FAILED(shell_link.As(&persist_file)))
    return;
  if (shortcut && FAILED(persist_file->Load(shortcut, STGM_READWRITE)))
    return;
  i_shell_link->Swap(shell_link);
  i_persist_file->Swap(persist_file);
}

} // namespace

bool UpdateShortcutLink(LPCTSTR shortcut_path, const ShortcutProperties& properties)
{
  bool shortcut_existed = is_file_exists(shortcut_path);

  // Interfaces to the old shortcut when replacing an existing shortcut.
  ComPtr<IShellLink> old_i_shell_link;
  ComPtr<IPersistFile> old_i_persist_file;

  // Interfaces to the shortcut being created/updated.
  ComPtr<IShellLink> i_shell_link;
  ComPtr<IPersistFile> i_persist_file;
  InitializeShortcutInterfaces(shortcut_path, &i_shell_link, &i_persist_file);

  // Return false immediately upon failure to initialize shortcut interfaces.
  if (!i_persist_file.Get())
    return false;

  if ((properties.options & ShortcutProperties::PROPERTIES_TARGET) && FAILED(i_shell_link->SetPath(properties.target)))
  {
    return false;
  }

  if ((properties.options & ShortcutProperties::PROPERTIES_WORKING_DIR) && FAILED(i_shell_link->SetWorkingDirectory(properties.working_dir)))
  {
    return false;
  }

  if (properties.options & ShortcutProperties::PROPERTIES_ARGUMENTS)
  {
    if (FAILED(i_shell_link->SetArguments(properties.arguments)))
      return false;
  }
  else if (old_i_persist_file.Get())
  {
    wchar_t current_arguments[MAX_PATH] = {0};
    if (SUCCEEDED(old_i_shell_link->GetArguments(current_arguments, MAX_PATH)))
    {
      i_shell_link->SetArguments(current_arguments);
    }
  }

  if ((properties.options & ShortcutProperties::PROPERTIES_DESCRIPTION) && FAILED(i_shell_link->SetDescription(properties.description)))
  {
    return false;
  }

  if ((properties.options & ShortcutProperties::PROPERTIES_ICON) &&
      FAILED(i_shell_link->SetIconLocation(properties.icon, properties.icon_index)))
  {
    return false;
  }

  bool has_app_id                = (properties.options & ShortcutProperties::PROPERTIES_APP_ID) != 0;
  //bool has_dual_mode             = (properties.options & ShortcutProperties::PROPERTIES_DUAL_MODE) != 0;
  //bool has_toast_activator_clsid = (properties.options & ShortcutProperties::PROPERTIES_TOAST_ACTIVATOR_CLSID) != 0;
  if (has_app_id/* || has_dual_mode || has_toast_activator_clsid*/)
  {
    ComPtr<IPropertyStore> property_store;
    if (FAILED(i_shell_link.As(&property_store)) || !property_store.Get())
      return false;

    if (has_app_id && !SetAppIdForPropertyStore(property_store.Get(), properties.app_id))
    {
      return false;
    }
    //if (has_dual_mode && !SetBooleanValueForPropertyStore(property_store.Get(), PKEY_AppUserModel_IsDualMode, properties.dual_mode))
    //{
    //  return false;
    //}
    //if (has_toast_activator_clsid && !SetClsidForPropertyStore(property_store.Get(), PKEY_AppUserModel_ToastActivatorCLSID, properties.toast_activator_clsid))
    //{
    //  return false;
    //}
  }

  // Release the interfaces to the old shortcut to make sure it doesn't prevent
  // overwriting it if needed.
  old_i_persist_file.Reset();
  old_i_shell_link.Reset();

  HRESULT result = i_persist_file->Save(shortcut_path, TRUE);

  // Release the interfaces in case the SHChangeNotify call below depends on
  // the operations above being fully completed.
  i_persist_file.Reset();
  i_shell_link.Reset();

  // If we successfully created/updated the icon, notify the shell that we have
  // done so.
  const bool succeeded = SUCCEEDED(result);
  if (succeeded)
  {
    if (shortcut_existed)
    {
      // TODO(gab): SHCNE_UPDATEITEM might be sufficient here; further testing
      // required.
      SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
    }
    else
    {
      SHChangeNotify(SHCNE_CREATE, SHCNF_PATH, shortcut_path, nullptr);
    }
  }

  return succeeded;
}

} // namespace win
} // namespace base


extern "C" BOOL SetShortcutAppId(LPCTSTR shortcut, LPCTSTR appid) {
	::CoInitialize(NULL);
	base::win::ShortcutProperties props;
	memset(&props, 0, sizeof(props));
	props.set_app_id(appid);
	BOOL ok = base::win::UpdateShortcutLink(shortcut, props);
	::CoUninitialize();
	return ok;
}
