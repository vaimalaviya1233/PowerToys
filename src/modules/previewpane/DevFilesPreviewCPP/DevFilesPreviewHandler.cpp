#include "pch.h"
#include "DevFilesPreviewHandler.h"
#include "Generated Files/resource.h"
#include "Settings.h"

#include <algorithm>
#include <filesystem>
#include <Shlwapi.h>
#include <wrl.h>

#include <common/SettingsAPI/settings_helpers.h>
#include <common/utils/gpo.h>
#include <common/utils/json.h>
#include <common/utils/process_path.h>
#include <common/utils/resources.h>
#include <common/Themes/theme_helpers.h>
#include <common/Themes/windows_colors.h>

using namespace Microsoft::WRL;

extern HINSTANCE    g_hInst;
extern long         g_cDllRef;

static const uint32_t cInfoBarHeight = 70;
static const size_t cMaxFileSize = 50000; // =~50KB
static const std::wstring cVirtualHostName = L"PowerToysLocalDevFilesPreview";

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

std::wstring get_file_content(std::wifstream& inStream, size_t fileSize)
{
    if (inStream)
    {
        std::wstring contents;
        contents.resize(fileSize);
        inStream.seekg(0, std::ios::beg);
        inStream.read(&contents[0], contents.size());
        inStream.close();
        return (contents);
    }
    throw(errno);
}

static const std::wstring base64_chars =
    L"ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "abcdefghijklmnopqrstuvwxyz"
    "0123456789+/";

static inline bool is_base64(wchar_t c)
{
    return (isalnum(c) || (c == '+') || (c == '/'));
}

std::wstring base64_encode(wchar_t const* buf, unsigned int bufLen)
{
    std::wstring ret;
    int i = 0;
    int j = 0;
    wchar_t char_array_3[3];
    wchar_t char_array_4[4];

    while (bufLen--)
    {
        char_array_3[i++] = *(buf++);
        if (i == 3)
        {
            char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
            char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
            char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
            char_array_4[3] = char_array_3[2] & 0x3f;

            for (i = 0; (i < 4); i++)
                ret += base64_chars[char_array_4[i]];
            i = 0;
        }
    }

    if (i)
    {
        for (j = i; j < 3; j++)
            char_array_3[j] = '\0';

        char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
        char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
        char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
        char_array_4[3] = char_array_3[2] & 0x3f;

        for (j = 0; (j < i + 1); j++)
            ret += base64_chars[char_array_4[j]];

        while ((i++ < 3))
            ret += '=';
    }

    return ret;
}
}

DevFilesPreviewHandler::DevFilesPreviewHandler() :
    m_cRef(1), m_hwndParent(NULL), m_rcParent(), m_punkSite(NULL), m_gpoText(NULL), m_errorText(NULL), m_infoBarAdded(false)
{
    m_webVew2UserDataFolder = PTSettingsHelper::get_local_low_folder_location() + L"\\DevFilesPreview-Temp";

    InterlockedIncrement(&g_cDllRef);
}

DevFilesPreviewHandler::~DevFilesPreviewHandler()
{
    if (m_gpoText)
    {
        DestroyWindow(m_gpoText);
        m_gpoText = NULL;
    }
    if (m_errorText)
    {
        DestroyWindow(m_errorText);
        m_errorText = NULL;
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
        if (m_errorText)
        {
            SetWindowPos(m_errorText, NULL, m_rcParent.left, m_rcParent.top,
                RECTWIDTH(m_rcParent), cInfoBarHeight,
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        if (m_webviewController)
        {
            if (!m_infoBarAdded)
            {
                m_webviewController->put_Bounds(m_rcParent);
            }
            else
            {
                RECT webViewRect{ m_rcParent.left, m_rcParent.top, m_rcParent.right, m_rcParent.bottom };
                webViewRect.top += cInfoBarHeight;

                m_webviewController->put_Bounds(webViewRect);
            }
        }
        hr = S_OK;

    }
    return hr;
}

IFACEMETHODIMP DevFilesPreviewHandler::DoPreview() {
    CleanupWebView2UserDataFolder();

    if (powertoys_gpo::getConfiguredMonacoPreviewEnabledValue() == powertoys_gpo::gpo_rule_configured_disabled)
    {
        m_infoBarAdded = true;
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

    std::wifstream in(m_filePath.c_str(), std::ios::in | std::ios::binary);
    in.seekg(0, std::ios::end);

    size_t fileSize = in.tellg();
    if (fileSize < cMaxFileSize)
    {
        std::wstring fileContent = get_file_content(in, fileSize);

        std::wstring vsCodeLangSet = GetLanguage(m_filePath);

        std::wifstream hmtlIn((get_module_folderpath(g_hInst) + L"\\index.html").c_str(), std::ios::in | std::ios::binary);
        hmtlIn.seekg(0, std::ios::end);
        fileSize = hmtlIn.tellg();
        std::wstring htmlFileContent = get_file_content(hmtlIn, fileSize);

        const std::wstring ptLang = L"[[PT_LANG]]";
        size_t pos = htmlFileContent.find(ptLang);
        if (pos != std::wstring::npos)
        {
            htmlFileContent.replace(pos, ptLang.size(), vsCodeLangSet);
        }

        const std::wstring ptWrap = L"[[PT_WRAP]]";
        pos = htmlFileContent.find(ptWrap);
        if (pos != std::wstring::npos)
        {
            auto& settingsInstance = DevFilesPreviewHandlerSettingsInstance();
            htmlFileContent = htmlFileContent.replace(pos, ptWrap.size(), settingsInstance.GetWrapText() ? L"1" : L"0");
        }

        bool light_mode = WindowsColors::rgb_color(WindowsColors::get_background_color()) != 0;

        const std::wstring ptTheme = L"[[PT_THEME]]";
        pos = htmlFileContent.find(ptTheme);
        if (pos != std::wstring::npos)
        {
            htmlFileContent = htmlFileContent.replace(pos, ptTheme.size(), light_mode ? L"light" : L"dark");
        }

        std::wstring encoded = base64_encode(fileContent.c_str(), (unsigned int)fileContent.size());

        const std::wstring ptCode = L"[[PT_CODE]]";
        pos = htmlFileContent.find(ptCode);
        if (pos != std::wstring::npos)
        {
            htmlFileContent = htmlFileContent.replace(pos, ptCode.size(), encoded);
        }

        const std::wstring ptUrl = L"[[PT_URL]]";
        pos = htmlFileContent.find(ptUrl);
        while (pos != std::wstring::npos)
        {
            htmlFileContent = htmlFileContent.replace(pos, ptUrl.size(), cVirtualHostName);
            pos = htmlFileContent.find(ptUrl, pos + ptUrl.size());
        }

        CreateCoreWebView2EnvironmentWithOptions(nullptr, m_webVew2UserDataFolder.c_str(), nullptr,
            Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>([this, htmlFileContent](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {
                // Create a CoreWebView2Controller and get the associated CoreWebView2 whose parent is the main window hWnd
                env->CreateCoreWebView2Controller(m_hwndParent, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>([this, htmlFileContent](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
                    if (controller != nullptr)
                    {
                        m_webviewController = controller;
                        m_webviewController->get_CoreWebView2(&m_webviewWindow);
                        wil::com_ptr<ICoreWebView2_3> webView;
                        webView = m_webviewWindow.try_query<ICoreWebView2_3>();
                        if (webView)
                        {
                            std::wstring modulePath = get_module_folderpath(g_hInst);
                            webView->SetVirtualHostNameToFolderMapping(cVirtualHostName.c_str(), modulePath.c_str(), COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_ALLOW);
                        }
                        else
                        {
                            PreviewError(L"Error setting virtual host");
                        }
                    }
                    else
                    {
                        PreviewError(L"Error initalizing WebView controller");
                    }

                    // Add a few settings for the webview
                    // The demo step is redundant since the values are the default settings
                    wil::com_ptr<ICoreWebView2Settings> Settings;
                    m_webviewWindow->get_Settings(&Settings);

                    Settings->put_AreDefaultScriptDialogsEnabled(FALSE);
                    Settings->put_AreDefaultContextMenusEnabled(FALSE);
                    Settings->put_IsScriptEnabled(TRUE);
                    Settings->put_IsZoomControlEnabled(FALSE);
                    Settings->put_IsStatusBarEnabled(FALSE);
                    Settings->put_IsBuiltInErrorPageEnabled(FALSE);
                    Settings->put_IsWebMessageEnabled(FALSE);
                    Settings->put_AreDevToolsEnabled(FALSE);
                    Settings->put_AreHostObjectsAllowed(FALSE);

                    wil::com_ptr<ICoreWebView2Settings4> webView2Settingsv4;
                    webView2Settingsv4 = Settings.try_query<ICoreWebView2Settings4>();
                    if (webView2Settingsv4)
                    {
                        webView2Settingsv4->put_IsGeneralAutofillEnabled(FALSE);
                        webView2Settingsv4->put_IsPasswordAutosaveEnabled(FALSE);
                    }

                    RECT bounds;
                    GetClientRect(m_hwndParent, &bounds);

                    if (!m_infoBarAdded)
                    {
                        m_webviewController->put_Bounds(bounds);                
                    }
                    else
                    {
                        bounds.top += cInfoBarHeight;
                        m_webviewController->put_Bounds(bounds);
                    }

                    m_webviewWindow->NavigateToString(htmlFileContent.c_str());

                    return S_OK;
                }).Get());
            return S_OK;
        }).Get());
    }
    else
    {
        PreviewError(GET_RESOURCE_STRING(IDS_MAXFILESIZEERROR));
    }

    return S_OK;
}

IFACEMETHODIMP DevFilesPreviewHandler::Unload()
{
    m_infoBarAdded = false;

    if (m_webviewController)
        m_webviewController->Close();
    if (m_webviewWindow)
        m_webviewWindow->Stop();
    if (m_gpoText)
    {
        DestroyWindow(m_gpoText);
        m_gpoText = NULL;
    }
    if (m_errorText)
    {
        DestroyWindow(m_errorText);
        m_errorText = NULL;
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


void DevFilesPreviewHandler::PreviewError(std::wstring errorMessage)
{
    m_infoBarAdded = true;

    if (m_errorText)
    {
        DestroyWindow(m_errorText);
    }

    m_errorText = CreateWindowEx(
        0, L"EDIT", // predefined class
        errorMessage.c_str(),
        WS_CHILD | WS_VISIBLE | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
        0,
        0,
        RECTWIDTH(m_rcParent),
        cInfoBarHeight,
        m_hwndParent, // parent window
        NULL, // edit control ID
        g_hInst,
        NULL);
}

void DevFilesPreviewHandler::CleanupWebView2UserDataFolder()
{
    try
    {
        // Cleanup temp
        for (auto const& dir_entry : std::filesystem::directory_iterator{ m_webVew2UserDataFolder })
        {
            if (dir_entry.path().extension().compare(L".html") == 0)
            {
                std::filesystem::remove(dir_entry);
            }
        }
    }
    catch (std::exception ex)
    {
        PreviewError(std::wstring{ L"Error cleaning up WebView2 user data folder:" + winrt::to_hstring(ex.what()) });

        // Trace::SvgFilePreviewError(ex.what());
    }
}

#pragma endregion