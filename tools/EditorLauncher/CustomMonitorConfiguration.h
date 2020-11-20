#pragma once

#include <string>

#include <winrt/Windows.Data.Json.h>

namespace CustomMonitorConfiguration
{
    struct MonitorConfig
    {
        int width;
        int height;
        int left;
        int top;
        int dpi;
        bool start;
    };

    void GenerateExample(const std::wstring& fileName);
    std::vector<MonitorConfig> ReadConfiguration(const char* fileName);

    std::wstring GenerateEditorArguments(const std::vector<MonitorConfig>& config);
}