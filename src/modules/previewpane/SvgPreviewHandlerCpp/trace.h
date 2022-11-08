#pragma once

#include "pch.h"

class Trace
{
public:
    static void RegisterProvider();
    static void UnregisterProvider();

    static void SvgFileHandlerLoaded();
    static void SvgFilePreviewed();
    static void SvgFilePreviewError(const char* exceptionMessage);
};
