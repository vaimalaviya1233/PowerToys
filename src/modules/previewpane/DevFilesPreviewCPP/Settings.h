#pragma once

#include "pch.h"

#include <string>

class DevFilesPreviewHandlerSettings
{
public:
    DevFilesPreviewHandlerSettings();

    inline bool GetWrapText()
    {
        Load();
        return m_wrapText;
    }

private:
    void Load();

    std::wstring m_jsonFilePath;

    bool m_wrapText;
};

DevFilesPreviewHandlerSettings& DevFilesPreviewHandlerSettingsInstance();
 
