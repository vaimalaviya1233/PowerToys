#include "pch.h"
#include "Settings.h"

#include <common/SettingsAPI/settings_helpers.h>
#include <common/utils/json.h>

static const std::wstring ModuleKey = L"File Explorer";

namespace
{
    const wchar_t c_properties[] = L"properties";
    const wchar_t c_textWrapJson[] = L"monaco-previewer-toggle-setting-word-wrap";
}

DevFilesPreviewHandlerSettings::DevFilesPreviewHandlerSettings() :
    m_wrapText(true)
{
    std::wstring result = PTSettingsHelper::get_module_save_folder_location(ModuleKey);
    m_jsonFilePath = result + std::wstring(L"\\settings.json");

    Load();
}

void DevFilesPreviewHandlerSettings::Load()
{
    auto json = json::from_file(m_jsonFilePath);
    if (json)
    {
        const json::JsonObject& jsonSettings = json.value();
        try
        {
            if (json::has(jsonSettings, c_properties))
            {
                auto properties = jsonSettings.GetNamedObject(L"properties", json::JsonObject{});
                if (json::has(properties, c_textWrapJson))
                {
                    m_wrapText = properties.GetNamedObject(c_textWrapJson).GetNamedBoolean(L"value");
                }
            }
        }
        catch (const winrt::hresult_error&)
        {
        }
    }
}

DevFilesPreviewHandlerSettings& DevFilesPreviewHandlerSettingsInstance()
{
    static DevFilesPreviewHandlerSettings instance;
    return instance;
}