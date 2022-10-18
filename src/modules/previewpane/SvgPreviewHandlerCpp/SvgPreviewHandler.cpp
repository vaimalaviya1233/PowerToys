#include "pch.h"
#include "SvgPreviewHandler.h"

#include <assert.h>
#include <exception>
#include <filesystem>
#include <Shlwapi.h>
#include <string>
#include <vector>
#include <wrl.h>

#include <WebView2.h>
#include <wil/com.h>

using namespace Microsoft::WRL;

extern HINSTANCE g_hInst;
extern long g_cDllRef;

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

}

SvgPreviewHandler::SvgPreviewHandler() :
    m_cRef(1), m_pStream(NULL), m_hwndParent(NULL), m_rcParent(), m_hwndPreview(NULL), m_punkSite(NULL)
{
    m_webVew2UserDataFolder = get_power_toys_local_low_folder_location() + L"\\SvgPreview-Temp";

    InterlockedIncrement(&g_cDllRef);
}

SvgPreviewHandler::~SvgPreviewHandler()
{
    if (m_hwndParent)
    {
        DestroyWindow(m_hwndPreview);
        m_hwndPreview = NULL;
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

IFACEMETHODIMP SvgPreviewHandler::SetFocus()
{
    HRESULT hr = S_FALSE;
    if (m_hwndPreview)
    {
        ::SetFocus(m_hwndPreview);
        hr = S_OK;
    }
    return hr;
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

IFACEMETHODIMP SvgPreviewHandler::DoPreview()
{
    HRESULT hr = E_FAIL;
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
        //PreviewError(ex, dataSource);
        //return;
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
        //PowerToysTelemetry.Log.WriteEvent(new SvgFilePreviewError{ Message = ex.Message });
    }

    //try
    //{
    //    _infoBarAdded = false;

    //    // Add a info bar on top of the Preview if any blocked element is present.
    //    if (blocked)
    //    {
    //        _infoBarAdded = true;
    //        AddTextBoxControl(Properties.Resource.BlockedElementInfoText);
    //    }

    AddWebViewControl(wsvgData);
    //    Resize += FormResized;
    //    base.DoPreview(dataSource);
    //    PowerToysTelemetry.Log.WriteEvent(new SvgFilePreviewed());
    //}
    //catch (Exception ex)
    //{
    //    PreviewError(ex, dataSource);
    //}

    return S_OK;
}

IFACEMETHODIMP SvgPreviewHandler::Unload()
{
    if (m_pStream)
    {
        m_pStream->Release();
        m_pStream = NULL;
    }

    if (m_hwndPreview)
    {
        DestroyWindow(m_hwndPreview);
        m_hwndPreview = NULL;
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

    CreateCoreWebView2EnvironmentWithOptions(nullptr, L"C:\\Users\\stefa\\AppData\\LocalLow\\Microsoft\\PowerToys\\aaaa", nullptr,
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>([this, svgData](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {
            // Create a CoreWebView2Controller and get the associated CoreWebView2 whose parent is the main window hWnd
            env->CreateCoreWebView2Controller(m_hwndParent, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>([this, svgData](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
                if (controller != nullptr)
                {
                    m_webviewController = controller;
                    m_webviewController->get_CoreWebView2(&m_webviewWindow);
                }

                // Add a few settings for the webview
                // The demo step is redundant since the values are the default settings
                ICoreWebView2Settings* Settings;
                m_webviewWindow->get_Settings(&Settings);
                //Settings->put_IsScriptEnabled(TRUE);
                //Settings->put_AreDefaultScriptDialogsEnabled(TRUE);
                //Settings->put_IsWebMessageEnabled(TRUE);

                // Resize WebView to fit the bounds of the parent window
                RECT bounds;
                GetClientRect(m_hwndParent, &bounds);
                m_webviewController->put_Bounds(bounds);

                // Schedule an async task to navigate to Bing
                //webviewWindow->Navigate(L"https://www.bing.com");

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
    try
    {
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
    }
    catch (std::exception)
    {
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
    catch (std::exception)
    {
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

#pragma endregion