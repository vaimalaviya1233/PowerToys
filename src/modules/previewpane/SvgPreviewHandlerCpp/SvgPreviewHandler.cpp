#include "pch.h"
#include "SvgPreviewHandler.h"
#include "Generated Files/resource.h"
#include "trace.h"

#include <assert.h>
#include <exception>
#include <filesystem>
#include <Shlwapi.h>
#include <string>
#include <vector>
#include <wrl.h>

#include <WebView2.h>
#include <wil/com.h>

#include <common/SettingsAPI/settings_helpers.h>
#include <common/utils/gpo.h>
#include <common/utils/resources.h>

using namespace Microsoft::WRL;

extern HINSTANCE g_hInst;
extern long g_cDllRef;

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

SvgPreviewHandler::SvgPreviewHandler() :
    m_cRef(1), m_pStream(NULL), m_hwndParent(NULL), m_rcParent(), m_punkSite(NULL), m_gpoText(NULL), m_infoBarAdded(false)
{
    m_webVew2UserDataFolder = PTSettingsHelper::get_local_low_folder_location() + L"\\SvgPreview-Temp";
    m_instance = this; 

    InterlockedIncrement(&g_cDllRef);
}

SvgPreviewHandler::~SvgPreviewHandler()
{
    m_instance = NULL; 
    if (m_gpoText)
    {
        DestroyWindow(m_gpoText);
        m_gpoText = NULL;
    }
    if (m_blockedText)
    {
        DestroyWindow(m_blockedText);
        m_blockedText = NULL;
    }
    if (m_errorText)
    {
        DestroyWindow(m_errorText);
        m_errorText = NULL;
    }
    if (m_punkSite)
    {
        m_punkSite->Release();
        m_punkSite = NULL;
    }
    if (m_pStream)
    {
        m_pStream->Release();
        m_pStream = NULL;
    }

    InterlockedDecrement(&g_cDllRef);
}

#pragma region IUnknown

IFACEMETHODIMP SvgPreviewHandler::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = {
        QITABENT(SvgPreviewHandler, IPreviewHandler),
        QITABENT(SvgPreviewHandler, IInitializeWithStream),
        QITABENT(SvgPreviewHandler, IPreviewHandlerVisuals),
        QITABENT(SvgPreviewHandler, IOleWindow),
        QITABENT(SvgPreviewHandler, IObjectWithSite),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) SvgPreviewHandler::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) SvgPreviewHandler::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion

#pragma region IInitializeWithStream

IFACEMETHODIMP SvgPreviewHandler::Initialize(IStream *pStream, DWORD grfMode)
{
    HRESULT hr = E_INVALIDARG;
    if (pStream)
    {
        if (m_pStream)
        {
            m_pStream->Release();
            m_pStream = NULL;
        }

        m_pStream = pStream;
        m_pStream->AddRef();
        hr = S_OK;

        Trace::SvgFileHandlerLoaded();
    }
    return hr;
}

#pragma endregion

#pragma region IPreviewHandler

IFACEMETHODIMP SvgPreviewHandler::SetWindow(HWND hwnd, const RECT *prc)
{
    if (hwnd && prc)
    {
        m_hwndParent = hwnd;
        m_rcParent = *prc;

    }
    return S_OK;
}

IFACEMETHODIMP SvgPreviewHandler::SetFocus()
{
    return S_OK;
}

IFACEMETHODIMP SvgPreviewHandler::QueryFocus(HWND *phwnd)
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

IFACEMETHODIMP SvgPreviewHandler::TranslateAccelerator(MSG *pmsg)
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

IFACEMETHODIMP SvgPreviewHandler::SetRect(const RECT *prc)
{
    HRESULT hr = E_INVALIDARG;
    if (prc != NULL)
    {
        m_rcParent = *prc;
        if (m_blockedText)
        {
            SetWindowPos(m_blockedText, NULL, m_rcParent.left, m_rcParent.top,
                RECTWIDTH(m_rcParent), cInfoBarHeight,
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
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

IFACEMETHODIMP SvgPreviewHandler::DoPreview()
{
    HRESULT hr = E_FAIL;

    if (powertoys_gpo::getConfiguredSvgPreviewEnabledValue() == powertoys_gpo::gpo_rule_configured_disabled)
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

    CleanupWebView2UserDataFolder();

    std::string svgData;
    std::wstring wsvgData;
    bool blocked = false;

    try
    {
        char buffer[4096];
        ULONG asd;
        while (true)
        {
            auto result = m_pStream->Read(buffer, 4096, &asd);
            svgData.append(buffer, asd);
            if (result == S_FALSE)
            {
                break;
            }
        }

        wsvgData = std::wstring{ winrt::to_hstring(svgData) };
        blocked = CheckBlockedElements(wsvgData);
    }
    catch (std::exception ex)
    {
        PreviewError(std::string("Error reading data:") + ex.what());

        Trace::SvgFilePreviewError(ex.what());
        return S_FALSE;
    }

    try
    {
        // Fixes #17527 - Inkscape v1.1 swapped order of default and svg namespaces in svg file (default first, svg after).
        // That resulted in parser being unable to parse it correctly and instead of svg, text was previewed.
        // MS Edge and Firefox also couldn't preview svg files with mentioned order of namespaces definitions.
        wsvgData = SwapNamespaces(wsvgData);
        wsvgData = AddStyleSVG(wsvgData);
    }
    catch (std::exception ex)
    {
        PreviewError(std::string("Error processing SVG data:") + ex.what());

        Trace::SvgFilePreviewError(ex.what());
    }

    try
    {
        m_infoBarAdded = false;

        // Add a info bar on top of the Preview if any blocked element is present.
        if (blocked)
        {
            m_infoBarAdded = true;

            m_blockedText = CreateWindowEx(
                0, L"EDIT", // predefined class
                GET_RESOURCE_STRING(IDS_BLOCKEDELEMENTINFOTEXT).c_str(),
                WS_CHILD | WS_VISIBLE | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                0,
                0,
                RECTWIDTH(m_rcParent),
                cInfoBarHeight,
                m_hwndParent, // parent window
                NULL, // edit control ID
                g_hInst,
                NULL);   

            UpdateWindow(m_hwndParent);
            UpdateWindow(m_blockedText);
        }
        AddWebViewControl(wsvgData);

        Trace::SvgFilePreviewed();
    }
    catch (std::exception ex)
    {
        PreviewError(std::string("Error previewing SVG data:") + ex.what());

        Trace::SvgFilePreviewError(ex.what());
    }

    return S_OK;
}

IFACEMETHODIMP SvgPreviewHandler::Unload()
{
    m_instance = NULL;
    if (m_pStream)
    {
        m_pStream->Release();
        m_pStream = NULL;
    }

    if (m_gpoText)
    {
        DestroyWindow(m_gpoText);
        m_gpoText = NULL;
    }
    if (m_blockedText)
    {
        DestroyWindow(m_blockedText);
        m_blockedText = NULL;
    }
    return S_OK;
}

#pragma endregion

#pragma region IPreviewHandlerVisuals

IFACEMETHODIMP SvgPreviewHandler::SetBackgroundColor(COLORREF color)
{
    return S_OK;
}

IFACEMETHODIMP SvgPreviewHandler::SetFont(const LOGFONTW *plf)
{
    return S_OK;
}

IFACEMETHODIMP SvgPreviewHandler::SetTextColor(COLORREF color)
{
    return S_OK;
}

#pragma endregion

#pragma region IOleWindow

IFACEMETHODIMP SvgPreviewHandler::GetWindow(HWND *phwnd)
{
    HRESULT hr = E_INVALIDARG;
    if (phwnd)
    {
        *phwnd = m_hwndParent;
        hr = S_OK;
    }
    return hr;
}

IFACEMETHODIMP SvgPreviewHandler::ContextSensitiveHelp(BOOL fEnterMode)
{
    return E_NOTIMPL;
}

#pragma endregion

#pragma region IObjectWithSite

IFACEMETHODIMP SvgPreviewHandler::SetSite(IUnknown *punkSite)
{
    if (m_punkSite)
    {
        m_punkSite->Release();
        m_punkSite = NULL;
    }
    return punkSite ? punkSite->QueryInterface(&m_punkSite) : S_OK;
}

IFACEMETHODIMP SvgPreviewHandler::GetSite(REFIID riid, void **ppv)
{
    *ppv = NULL;
    return m_punkSite ? m_punkSite->QueryInterface(riid, ppv) : E_FAIL;
}

#pragma endregion

#pragma region Helper Functions

void SvgPreviewHandler::AddWebViewControl(std::wstring svgData)
{
    if (m_webviewController)
        m_webviewController->Close();
    if (m_webviewWindow)
        m_webviewWindow->Stop();
    CreateCoreWebView2EnvironmentWithOptions(nullptr, m_webVew2UserDataFolder.c_str(), nullptr,
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>([this, svgData](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {
            // Create a CoreWebView2Controller and get the associated CoreWebView2 whose parent is the main window hWnd
            env->CreateCoreWebView2Controller(m_hwndParent, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>([this, svgData](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
                if (controller != nullptr)
                {
                    m_webviewController = controller;
                    m_webviewController->get_CoreWebView2(&m_webviewWindow);
                }
                else
                {
                    PreviewError(std::string("Error initalizing WebView controller"));
                }

                // Add a few settings for the webview
                // The demo step is redundant since the values are the default settings
                ICoreWebView2Settings* Settings;
                m_webviewWindow->get_Settings(&Settings);

                Settings->put_AreDefaultScriptDialogsEnabled(FALSE);
                Settings->put_AreDefaultContextMenusEnabled(FALSE);
                Settings->put_IsScriptEnabled(FALSE);
                Settings->put_IsWebMessageEnabled(FALSE);
                Settings->put_AreDevToolsEnabled(FALSE);
                Settings->put_AreHostObjectsAllowed(FALSE);
                // TODO(stefan)
                //_browser.CoreWebView2.Settings.IsGeneralAutofillEnabled = false;
                //_browser.CoreWebView2.Settings.IsPasswordAutosaveEnabled = false;

                // TODO(stefan)
                //_browser.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
                //_browser.CoreWebView2.WebResourceRequested += CoreWebView2_BlockExternalResources;

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
                // TODO(stefan) Add 2MB check
                m_webviewWindow->NavigateToString(svgData.c_str());

                return S_OK;
            }).Get());
        return S_OK;
    }).Get());
}

BOOL SvgPreviewHandler::CheckBlockedElements(std::wstring svgData)
{
    const std::vector<std::wstring> BlockedElementsName = { L"script", L"image", L"feimage" };

    bool foundBlockedElement = false;
    if (svgData.empty())
    {
        return foundBlockedElement;
    }

    // Check if any of the blocked element is present. If failed to parse or iterate over Svg return default false.
    // No need to throw because all the external content and script are blocked on the Web Browser Control itself.
    winrt::Windows::Data::Xml::Dom::XmlDocument doc;
    doc.LoadXml(svgData);

    for (const auto blockedElem : BlockedElementsName)
    {
        winrt::Windows::Data::Xml::Dom::XmlNodeList elems = doc.GetElementsByTagName(blockedElem);
        if (elems.Size() > 0)
        {
            foundBlockedElement = true;
            break;
        }
    }

    return foundBlockedElement;
}

void SvgPreviewHandler::CleanupWebView2UserDataFolder()
{
    try
    {
        // Cleanup temp
        for (auto const& dir_entry : std::filesystem::directory_iterator{ m_webVew2UserDataFolder })
        {
            if (dir_entry.path().extension().compare(L"html") == 0)
            {
                std::filesystem::remove(dir_entry);
            }
        }
    }
    catch (std::exception ex)
    {
        PreviewError(std::string("Error cleaning up WebView2 user data folder:") + ex.what());

        Trace::SvgFilePreviewError(ex.what());
    }
}

std::wstring SvgPreviewHandler::SwapNamespaces(std::wstring svgData)
{
    const std::wstring defaultNamespace = L"xmlns=\"http://www.w3.org/2000/svg\"";
    const std::wstring svgNamespace = L"xmlns:svg=\"http://www.w3.org/2000/svg\"";

    size_t defaultNamespaceIndex = svgData.find(defaultNamespace);
    size_t svgNamespaceIndex = svgData.find(svgNamespace);

    if (defaultNamespaceIndex != std::wstring::npos && svgNamespaceIndex != std::wstring::npos && defaultNamespaceIndex < svgNamespaceIndex)
    {
        const std::wstring first{ L"{0}" };
        const std::wstring second{ L"{1}" };
        svgData = svgData.replace(defaultNamespaceIndex, defaultNamespace.length(), L"{0}");
        svgData = svgData.replace(svgNamespaceIndex, svgNamespace.length(), L"{1}");

        svgData = svgData.replace(defaultNamespaceIndex, first.length(), svgNamespace);
        svgData = svgData.replace(svgNamespaceIndex, second.length(), defaultNamespace);
    }

    return svgData;
}

std::wstring SvgPreviewHandler::AddStyleSVG(std::wstring svgData)
{
    winrt::Windows::Data::Xml::Dom::XmlDocument doc;
    doc.LoadXml(svgData);

    winrt::Windows::Data::Xml::Dom::XmlNodeList elems = doc.GetElementsByTagName(L"svg");

    std::wstring width;
    std::wstring height;
    std::wstring widthR;
    std::wstring heightR;
    std::wstring oldStyle;
    bool hasViewBox = false;

    if (elems.Size() > 0)
    {
        auto svgElem = *elems.First();
        auto attributes = svgElem.Attributes();

        for (uint32_t i = 0; i < attributes.Size(); ++i)
        {
            auto elem = attributes.GetAt(i);
            if (elem.NodeName() == L"height")
            {
                height = elem.NodeValue().try_as<winrt::hstring>().value();
                attributes.RemoveNamedItem(L"height");
                i--;
            }
            else if (elem.NodeName() == L"width")
            {
                width = elem.NodeValue().try_as<winrt::hstring>().value();
                attributes.RemoveNamedItem(L"width");
                i--;
            }
            else if (elem.NodeName() == L"style")
            {
                oldStyle = elem.NodeValue().try_as<winrt::hstring>().value();
                attributes.RemoveNamedItem(L"style");
                i--;
            }
            else if (elem.NodeName() == L"viewBox")
            {
                hasViewBox = true;
            }
        }

        height = CheckUnit(height);
        width = CheckUnit(width);
        heightR = RemoveUnit(height);
        widthR = RemoveUnit(width);

        std::wstring centering = L"position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);";

        // max-width and max-height not supported. Extra CSS is needed for it to work.
        std::wstring scaling = L"max-width: {" + width + L"} ; max-height: {" + height + L"} ;";
        scaling += L"  _height:expression(this.scrollHeight > {" + heightR + L"} ? \" {" + height + L"}\" : \"auto\"); _width:expression(this.scrollWidth > {" + widthR + L"} ? \"{" + width + L"}\" : \"auto\");";

        winrt::Windows::Data::Xml::Dom::XmlAttribute styleAttr = doc.CreateAttribute(L"style");
        styleAttr.Value(scaling + centering + oldStyle);
        attributes.SetNamedItem(styleAttr);

        if (!hasViewBox)
        {
            // Fixes https://github.com/microsoft/PowerToys/issues/18107
            std::wstring viewBox = L"0 0 {" + widthR + L"} {" + heightR + L"}";
            winrt::Windows::Data::Xml::Dom::XmlAttribute viewBoxAttr = doc.CreateAttribute(L"viewBox");
            viewBoxAttr.Value(viewBox);
            attributes.SetNamedItem(viewBoxAttr);
        }

        return std::wstring{ doc.GetXml() };
    }

    return svgData;
}

std::wstring SvgPreviewHandler::CheckUnit(std::wstring value)
{
    std::vector<std::wstring> cssUnits = { L"cm", L"mm", L"in", L"px", L"pt", L"pc", L"em", L"ex", L"ch", L"rem", L"vw", L"vh", L"vmin", L"vmax", L"%" };

    for (std::wstring unit : cssUnits)
    {
        if (value.ends_with(unit))
        {
            return value;
        }
    }

    return value + L"px";
}

std::wstring SvgPreviewHandler::RemoveUnit(std::wstring value)
{
    std::vector<std::wstring> cssUnits = { L"cm", L"mm", L"in", L"px", L"pt", L"pc", L"em", L"ex", L"ch", L"rem", L"vw", L"vh", L"vmin", L"vmax", L"%" };

    for (std::wstring unit : cssUnits)
    {
        if (value.ends_with(unit))
        {
            value = value.erase(value.size() - unit.size());
            return value;
        }
    }

    return value;
}

void SvgPreviewHandler::PreviewError(std::string errorMessage)
{
    if (m_errorText)
    {
        DestroyWindow(m_errorText);
    }

    m_errorText = CreateWindowEx(
        0, L"EDIT", // predefined class
        winrt::to_hstring(errorMessage).c_str(),
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

#pragma endregion
