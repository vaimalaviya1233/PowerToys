#include "pch.h"
#include "trace.h"

/*
*
* This file captures the telemetry for the File Explorer Custom Renders project.
* The following telemetry is to be captured for this library:
* (1.) Is the previewer enabled.  
* (2.) File rendered per user in 24 hrs per file time (one for MD, one for SVG)
* (3.) Crashes.
*
*/

TRACELOGGING_DEFINE_PROVIDER(
    g_hProvider,
    "Microsoft.PowerToys",
    // {38e8889b-9731-53f5-e901-e8a7c1753074}
    (0x38e8889b, 0x9731, 0x53f5, 0xe9, 0x01, 0xe8, 0xa7, 0xc1, 0x75, 0x30, 0x74),
    TraceLoggingOptionProjectTelemetry());

void Trace::RegisterProvider()
{
    TraceLoggingRegister(g_hProvider);
}

void Trace::UnregisterProvider()
{
    TraceLoggingUnregister(g_hProvider);
}

void Trace::SvgFileHandlerLoaded()
{
    TraceLoggingWrite(
        g_hProvider,
        "SvgPreviewHandler_Loaded",
        ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
        TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}

void Trace::SvgFilePreviewed()
{
    TraceLoggingWrite(
        g_hProvider,
        "SvgPreviewHandler_Previewed",
        ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
        TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}

void Trace::SvgFilePreviewError(const char* exceptionMessage)
{
    TraceLoggingWrite(
        g_hProvider,
        "SvgPreviewHandler_Error",
        TraceLoggingString(exceptionMessage, "ExceptionMessage"),
        ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
        TraceLoggingBoolean(TRUE, "UTCReplace_AppSessionGuid"),
        TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}
