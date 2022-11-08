#include "pch.h"
#include "DevFilesPreviewHandler.h"
#include "Generated Files/resource.h"

#include <algorithm>
#include <filesystem>
#include <Shlwapi.h>

#include <common/SettingsAPI/settings_helpers.h>
#include <common/utils/gpo.h>
#include <common/utils/json.h>
#include <common/utils/process_path.h>
#include <common/utils/resources.h>

extern HINSTANCE    g_hInst;
extern long         g_cDllRef;

static const uint32_t cInfoBarHeight = 70;

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

}

DevFilesPreviewHandler::DevFilesPreviewHandler() :
    m_cRef(1), m_hwndParent(NULL), m_rcParent(), m_punkSite(NULL), m_gpoText(NULL)
{
    InterlockedIncrement(&g_cDllRef);
}

DevFilesPreviewHandler::~DevFilesPreviewHandler()
{
    if (m_gpoText)
    {
        DestroyWindow(m_gpoText);
        m_gpoText = NULL;
    }

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
    }
    return S_OK;
}

IFACEMETHODIMP DevFilesPreviewHandler::SetFocus()
{
    return S_OK;
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
        hr = S_OK;
    }
    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::DoPreview() {
    HRESULT hr = E_FAIL;

    if (powertoys_gpo::getConfiguredMonacoPreviewEnabledValue() == powertoys_gpo::gpo_rule_configured_disabled)
    {
        // GPO is disabling this utility. Show an error message instead.
        m_gpoText = CreateWindowEx(
            0, L"EDIT", // predefined class
            GET_RESOURCE_STRING(IDS_GPODISABLEDERRORTEXT).c_str(),
            WS_CHILD | WS_VISIBLE | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
            5,
            5,
            RECTWIDTH(m_rcParent) - 10,
            cInfoBarHeight,
            m_hwndParent,
            NULL,
            g_hInst,
            NULL);

        return S_OK;
    }

    auto asd = GetLanguage(std::filesystem::path{m_filePath}.extension());
    MessageBox(NULL, asd.c_str(), L"AAAA", NULL);

    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::Unload()
{
    if (m_gpoText)
    {
        DestroyWindow(m_gpoText);
        m_gpoText = NULL;
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