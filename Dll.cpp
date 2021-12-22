// QOI Thumbnail Provider for Windows Explorer
// Written by iOrange in 2021
// 
// Based on Microsoft's example
// https://github.com/microsoft/windows-classic-samples/tree/main/Samples/Win7Samples/winui/shell/appshellintegration/RecipeThumbnailProvider
// 
// Also more info here:
// https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/cc144118(v=vs.85)

#include <objbase.h>
#include <shlwapi.h>
#include <thumbcache.h> // For IThumbnailProvider.
#include <shlobj.h>     // For SHChangeNotify
#include <new>
#include <atomic>
#include <vector>       // For std::size

#ifdef QOI_THUMB_DEBUG_OUTPUT_ENABLED
#define QOI_THUMB_DEBUG_OUTPUT(s) OutputDebugStringA()
#else
#define QOI_THUMB_DEBUG_OUTPUT(s)
#endif


extern HRESULT CQOIThumbProvider_CreateInstance(REFIID riid, void** ppv);

#define SZ_CLSID_QOITHUMBHANDLER    L"{98238d8e-7201-4588-bd77-61e41ad3e977}"
#define SZ_QOITHUMBHANDLER          L"QOI Thumbnail Handler"

constexpr CLSID kCLSID_QOIThumbHandler = { 0x98238d8e, 0x7201, 0x4588, { 0xbd, 0x77, 0x61, 0xe4, 0x1a, 0xd3, 0xe9, 0x77 } };

typedef HRESULT(*PFNCREATEINSTANCE)(REFIID riid, void** ppvObject);
struct CLASS_OBJECT_INIT {
    const CLSID*        pClsid;
    PFNCREATEINSTANCE   pfnCreate;
};

// add classes supported by this module here
constexpr CLASS_OBJECT_INIT kClassObjectInit[] = {
    { &kCLSID_QOIThumbHandler, CQOIThumbProvider_CreateInstance }
};


std::atomic_long    gModuleReferences(0);
HINSTANCE           gModuleInstance = nullptr;

// Standard DLL functions
STDAPI_(BOOL) DllMain(HINSTANCE hInstance, DWORD dwReason, void*) {
    if (DLL_PROCESS_ATTACH == dwReason) {
        gModuleInstance = hInstance;
        ::DisableThreadLibraryCalls(hInstance);
    } else if (DLL_PROCESS_DETACH == dwReason) {
        gModuleInstance = nullptr;
    }
    return TRUE;
}

STDAPI DllCanUnloadNow() {
    // Only allow the DLL to be unloaded after all outstanding references have been released
    return (gModuleReferences > 0) ? S_FALSE : S_OK;
}

void DllAddRef() {
    ++gModuleReferences;
}

void DllRelease() {
    --gModuleReferences;
}

class CClassFactory : public IClassFactory {
public:
    static HRESULT CreateInstance(REFCLSID clsid, const CLASS_OBJECT_INIT* pClassObjectInits, size_t cClassObjectInits, REFIID riid, void** ppv) {
        *ppv = NULL;
        HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;
        for (size_t i = 0; i < cClassObjectInits; ++i) {
            if (clsid == *pClassObjectInits[i].pClsid) {
                IClassFactory* pClassFactory = new (std::nothrow) CClassFactory(pClassObjectInits[i].pfnCreate);
                hr = pClassFactory ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr)) {
                    hr = pClassFactory->QueryInterface(riid, ppv);
                    pClassFactory->Release();
                }
                break; // match found
            }
        }
        return hr;
    }

    CClassFactory(PFNCREATEINSTANCE pfnCreate)
        : mReferences(1)
        , mCreateFunc(pfnCreate) {
        DllAddRef();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) {
        static const QITAB qit[] = {
            QITABENT(CClassFactory, IClassFactory),
            { 0 }
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() {
        return ++mReferences;
    }

    IFACEMETHODIMP_(ULONG) Release() {
        const long refs = --mReferences;
        if (!refs) {
            delete this;
        }
        return refs;
    }

    // IClassFactory
    IFACEMETHODIMP CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv) {
        return punkOuter ? CLASS_E_NOAGGREGATION : mCreateFunc(riid, ppv);
    }

    IFACEMETHODIMP LockServer(BOOL fLock) {
        if (fLock) {
            DllAddRef();
        } else {
            DllRelease();
        }
        return S_OK;
    }

private:
    ~CClassFactory() {
        DllRelease();
    }

    std::atomic_long    mReferences;
    PFNCREATEINSTANCE   mCreateFunc;
};

STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv) {
    return CClassFactory::CreateInstance(clsid, kClassObjectInit, std::size(kClassObjectInit), riid, ppv);
}

// A struct to hold the information required for a registry entry
struct REGISTRY_ENTRY {
    HKEY   hkeyRoot;
    PCWSTR pszKeyName;
    PCWSTR pszValueName;
    PCWSTR pszData;
};

// Creates a registry key (if needed) and sets the default value of the key
HRESULT CreateRegKeyAndSetValue(const REGISTRY_ENTRY* pRegistryEntry) {
    HKEY hKey;
    HRESULT hr = HRESULT_FROM_WIN32(RegCreateKeyExW(pRegistryEntry->hkeyRoot,
                                                    pRegistryEntry->pszKeyName,
                                                    0, nullptr, REG_OPTION_NON_VOLATILE,
                                                    KEY_SET_VALUE | KEY_WOW64_64KEY,
                                                    nullptr, &hKey, nullptr));
    if (SUCCEEDED(hr)) {
        hr = HRESULT_FROM_WIN32(RegSetValueExW(hKey, pRegistryEntry->pszValueName, 0, REG_SZ,
                                               reinterpret_cast<const BYTE*>(pRegistryEntry->pszData),
                                               static_cast<DWORD>(wcslen(pRegistryEntry->pszData) + 1) * sizeof(WCHAR)));
        RegCloseKey(hKey);
    }
    return hr;
}

// Registers this COM server
STDAPI DllRegisterServer() {
    HRESULT hr;
    WCHAR szModuleName[MAX_PATH] = { 0 };

    if (!GetModuleFileNameW(gModuleInstance, szModuleName, ARRAYSIZE(szModuleName))) {
        hr = HRESULT_FROM_WIN32(GetLastError());
    } else {
        // List of registry entries we want to create
        const REGISTRY_ENTRY registryEntries[] = {
            // RootKey          KeyName                                                                      ValueName          Data
            {HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_QOITHUMBHANDLER,                      nullptr,           SZ_QOITHUMBHANDLER},
            {HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_QOITHUMBHANDLER L"\\InProcServer32",  nullptr,           szModuleName},
            {HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_QOITHUMBHANDLER L"\\InProcServer32",  L"ThreadingModel", L"Apartment"},
            {HKEY_CURRENT_USER, L"Software\\Classes\\.qoi\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}", nullptr,           SZ_CLSID_QOITHUMBHANDLER},
        };

        hr = S_OK;
        for (size_t i = 0; i < std::size(registryEntries) && SUCCEEDED(hr); ++i) {
            hr = CreateRegKeyAndSetValue(&registryEntries[i]);
        }
    }

    if (SUCCEEDED(hr)) {
        // This tells the shell to invalidate the thumbnail cache.  This is important because any .qoi files
        // viewed before registering this handler would otherwise show cached blank thumbnails.
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
    }

    return hr;
}

// Unregisters this COM server
STDAPI DllUnregisterServer() {
    HRESULT hr = S_OK;

    const PCWSTR regKeys[] = {
        L"Software\\Classes\\CLSID\\" SZ_CLSID_QOITHUMBHANDLER,
        L"Software\\Classes\\.qoi\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"
    };

    // Delete the registry entries
    for (size_t i = 0; i < std::size(regKeys) && SUCCEEDED(hr); ++i) {
        hr = HRESULT_FROM_WIN32(RegDeleteTreeW(HKEY_CURRENT_USER, regKeys[i]));
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) {
            // If the registry entry has already been deleted, say S_OK.
            hr = S_OK;
        }
    }

    return hr;
}
