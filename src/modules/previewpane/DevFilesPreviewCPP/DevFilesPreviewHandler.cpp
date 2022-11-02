#include "pch.h"
#include "DevFilesPreviewHandler.h"

#include <algorithm>
#include <filesystem>
#include <Shlwapi.h>

#include <common/utils/json.h>

extern HINSTANCE    g_hInst;
extern long         g_cDllRef;

namespace
{
inline int RECTWIDTH(const RECT& rc)
{
    return (rc.right - rc.left);
}

inline int RECTHEIGHT(const RECT& rc)
{
    return (rc.bottom - rc.top);
}

std::wstring get_power_toys_local_low_folder_location()
{
    PWSTR local_app_path;
    winrt::check_hresult(SHGetKnownFolderPath(FOLDERID_LocalAppDataLow, 0, NULL, &local_app_path));
    std::wstring result{ local_app_path };
    CoTaskMemFree(local_app_path);

    result += L"\\Microsoft\\PowerToys";
    std::filesystem::path save_path(result);
    if (!std::filesystem::exists(save_path))
    {
        std::filesystem::create_directories(save_path);
    }
    return result;
}

inline std::wstring get_module_folderpath(HMODULE mod = nullptr, const bool removeFilename = true)
{
    wchar_t buffer[MAX_PATH + 1];
    DWORD actual_length = GetModuleFileNameW(mod, buffer, MAX_PATH);
    if (GetLastError() == ERROR_INSUFFICIENT_BUFFER)
    {
        const DWORD long_path_length = 0xFFFF; // should be always enough
        std::wstring long_filename(long_path_length, L'\0');
        actual_length = GetModuleFileNameW(mod, long_filename.data(), long_path_length);
        PathRemoveFileSpecW(long_filename.data());
        long_filename.resize(std::wcslen(long_filename.data()));
        long_filename.shrink_to_fit();
        return long_filename;
    }

    if (removeFilename)
    {
        PathRemoveFileSpecW(buffer);
    }
    return { buffer, (UINT)lstrlenW(buffer) };
}

}

DevFilesPreviewHandler::DevFilesPreviewHandler() :
    m_cRef(1), m_hwndParent(NULL), m_rcParent(), m_hwndPreview(NULL), m_punkSite(NULL)
{
    InterlockedIncrement(&g_cDllRef);
}

DevFilesPreviewHandler::~DevFilesPreviewHandler()
{
    InterlockedDecrement(&g_cDllRef);
}

#pragma region IUnknown

IFACEMETHODIMP DevFilesPreviewHandler::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(DevFilesPreviewHandler, IPreviewHandler),
        QITABENT(DevFilesPreviewHandler, IInitializeWithFile),
        QITABENT(DevFilesPreviewHandler, IPreviewHandlerVisuals),
        QITABENT(DevFilesPreviewHandler, IOleWindow),
        QITABENT(DevFilesPreviewHandler, IObjectWithSite),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) DevFilesPreviewHandler::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) DevFilesPreviewHandler::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }
    return cRef;
}

#pragma endregion

#pragma region IInitializationWithFile

IFACEMETHODIMP DevFilesPreviewHandler::Initialize(LPCWSTR pszFilePath, DWORD grfMode)
{
    m_filePath = pszFilePath;
    return S_OK;
}

#pragma endregion

#pragma region IPreviewHandler

IFACEMETHODIMP DevFilesPreviewHandler::SetWindow(HWND hwnd, const RECT *prc)
{
    if (hwnd && prc)
    {
        m_hwndParent = hwnd;
        m_rcParent = *prc;

        if (m_hwndPreview)
        {
            SetParent(m_hwndPreview, m_hwndParent);
            SetWindowPos(m_hwndPreview, NULL, m_rcParent.left, m_rcParent.top,
                RECTWIDTH(m_rcParent), RECTHEIGHT(m_rcParent),
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }
    return S_OK;
}

IFACEMETHODIMP DevFilesPreviewHandler::SetFocus()
{
    HRESULT hr = S_FALSE;
    if (m_hwndPreview)
    {
        ::SetFocus(m_hwndPreview);
        hr = S_OK;
    }
    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::QueryFocus(HWND *phwnd)
{
    HRESULT hr = E_INVALIDARG;
    if (phwnd)
    {
        *phwnd = ::GetFocus();
        if (*phwnd)
        {
            hr = S_OK;
        }
        else
        {
            hr = HRESULT_FROM_WIN32(GetLastError());
        }
    }
    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::TranslateAccelerator(MSG *pmsg)
{
    HRESULT hr = S_FALSE;
    IPreviewHandlerFrame* pFrame = NULL;
    if (m_punkSite && SUCCEEDED(m_punkSite->QueryInterface(&pFrame)))
    {
        hr = pFrame->TranslateAccelerator(pmsg);

        pFrame->Release();
    }
    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::SetRect(const RECT *prc)
{
    HRESULT hr = E_INVALIDARG;
    if (prc != NULL)
    {
        m_rcParent = *prc;
        if (m_hwndPreview)
        {
            SetWindowPos(m_hwndPreview, NULL, m_rcParent.left, m_rcParent.top,
                RECTWIDTH(m_rcParent), RECTHEIGHT(m_rcParent),
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
        hr = S_OK;
    }
    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::DoPreview() {
    HRESULT hr = S_OK;
    auto asd = GetLanguage(std::filesystem::path{m_filePath}.extension());

    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::Unload()
{
    if (m_hwndPreview)
    {
        DestroyWindow(m_hwndPreview);
        m_hwndPreview = NULL;
    }
    return S_OK;
}

#pragma endregion

#pragma region IPreviewHandlerVisuals

IFACEMETHODIMP DevFilesPreviewHandler::SetBackgroundColor(COLORREF color)
{
    return S_OK;
}

IFACEMETHODIMP DevFilesPreviewHandler::SetFont(const LOGFONTW* plf)
{
    return S_OK;
}

IFACEMETHODIMP DevFilesPreviewHandler::SetTextColor(COLORREF color)
{
    return S_OK;
}

#pragma endregion

#pragma region IOleWindow

IFACEMETHODIMP DevFilesPreviewHandler::GetWindow(HWND* phwnd)
{
    HRESULT hr = E_INVALIDARG;
    if (phwnd)
    {
        *phwnd = m_hwndParent;
        hr = S_OK;
    }
    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::ContextSensitiveHelp(BOOL fEnterMode)
{
    return E_NOTIMPL;
}

#pragma endregion

#pragma region IObjectWithSite

IFACEMETHODIMP DevFilesPreviewHandler::SetSite(IUnknown* punkSite)
{
    if (m_punkSite)
    {
        m_punkSite->Release();
        m_punkSite = NULL;
    }
    return punkSite ? punkSite->QueryInterface(&m_punkSite) : S_OK;
}

IFACEMETHODIMP DevFilesPreviewHandler::GetSite(REFIID riid, void** ppv)
{
    *ppv = NULL;
    return m_punkSite ? m_punkSite->QueryInterface(riid, ppv) : E_FAIL;
}

#pragma endregion

#pragma region Helper Functions

std::wstring DevFilesPreviewHandler::GetLanguage(std::wstring fileExtension)
{
    std::transform(fileExtension.begin(), fileExtension.end(), fileExtension.begin(), ::towlower);
    
    auto languageListDocument = get_module_folderpath(g_hInst) + L"\\monaco_languages.json";
    auto json = json::from_file(languageListDocument);

    if (json)
    {
        try
        {
            auto list = json->GetNamedArray(L"list");
            for (uint32_t i = 0; i < list.Size(); ++i)
            {
                auto entry = list.GetObjectAt(i);
                auto extensionsList = entry.GetNamedArray(L"extensions");

                for (uint32_t j = 0; j < extensionsList.Size(); ++j)
                {
                    auto extension = extensionsList.GetStringAt(j);
                    if (extension == fileExtension)
                    {
                        return std::wstring{ entry.GetNamedString(L"id") };
                    }
                }
            }
        }
        catch (...)
        {
        }
    }
    return {};
}

#pragma endregion