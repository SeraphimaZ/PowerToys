#include "CustomMonitorConfiguration.h"
#include "pch.h"

#include <iostream>
#include <filesystem>
#include <fstream>

#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Data.Json.h>

#include <optional>

namespace json
{
    using namespace winrt::Windows::Data::Json;

    void to_file(std::wstring_view file_name, const JsonObject& obj)
    {
        std::wstring obj_str{ obj.Stringify().c_str() };
        std::ofstream{ file_name.data(), std::ios::binary } << winrt::to_string(obj_str);
    }

    JsonObject from_file(const char* fileName)
    {
        std::ifstream file(fileName, std::ios::binary);
        if (file.is_open())
        {
            using isbi = std::istreambuf_iterator<char>;
            std::wstring obj_str{ isbi{ file }, isbi{} };
            JsonObject obj(JsonObject::Parse(obj_str));
            return obj;
        }

        return JsonObject();
    }
}

namespace JsonTags
{
    const std::wstring MonitorsTag = L"monitors";
    const std::wstring WidthTag = L"width";
    const std::wstring HeightTag = L"height";
    const std::wstring LeftTag = L"left";
    const std::wstring TopTag = L"top";
    const std::wstring DpiTag = L"dpi";
    const std::wstring UseAsStartMonitorTag = L"useAsStartMonitor";
}

namespace CustomMonitorConfiguration
{
    void GenerateExample(const std::wstring& fileName)
    {
        json::JsonObject result{};

        json::JsonArray array{};

        json::JsonObject monitor1{};
        monitor1.SetNamedValue(JsonTags::WidthTag, json::JsonValue::CreateNumberValue(1920));
        monitor1.SetNamedValue(JsonTags::HeightTag, json::JsonValue::CreateNumberValue(1080));
        monitor1.SetNamedValue(JsonTags::LeftTag, json::JsonValue::CreateNumberValue(0));
        monitor1.SetNamedValue(JsonTags::TopTag, json::JsonValue::CreateNumberValue(0));
        monitor1.SetNamedValue(JsonTags::DpiTag, json::JsonValue::CreateNumberValue(96));
        monitor1.SetNamedValue(JsonTags::UseAsStartMonitorTag, json::JsonValue::CreateBooleanValue(true));
        array.Append(std::move(monitor1));

        json::JsonObject monitor2{};
        monitor2.SetNamedValue(JsonTags::WidthTag, json::JsonValue::CreateNumberValue(2880));
        monitor2.SetNamedValue(JsonTags::HeightTag, json::JsonValue::CreateNumberValue(1620));
        monitor2.SetNamedValue(JsonTags::LeftTag, json::JsonValue::CreateNumberValue(-2880));
        monitor2.SetNamedValue(JsonTags::TopTag, json::JsonValue::CreateNumberValue(-60));
        monitor2.SetNamedValue(JsonTags::DpiTag, json::JsonValue::CreateNumberValue(96));
        monitor2.SetNamedValue(JsonTags::UseAsStartMonitorTag, json::JsonValue::CreateBooleanValue(false));
        array.Append(std::move(monitor2));

        json::JsonObject monitor3{};
        monitor3.SetNamedValue(JsonTags::WidthTag, json::JsonValue::CreateNumberValue(1920));
        monitor3.SetNamedValue(JsonTags::HeightTag, json::JsonValue::CreateNumberValue(1080));
        monitor3.SetNamedValue(JsonTags::LeftTag, json::JsonValue::CreateNumberValue(1920));
        monitor3.SetNamedValue(JsonTags::TopTag, json::JsonValue::CreateNumberValue(0));
        monitor3.SetNamedValue(JsonTags::DpiTag, json::JsonValue::CreateNumberValue(144));
        monitor3.SetNamedValue(JsonTags::UseAsStartMonitorTag, json::JsonValue::CreateBooleanValue(false));
        array.Append(std::move(monitor3));

        result.SetNamedValue(JsonTags::MonitorsTag, array);
        json::to_file(fileName, result);
    }

    std::vector<MonitorConfig> ReadConfiguration(const char* fileName)
    {
        std::vector<MonitorConfig> result{};
        
        json::JsonObject jsonObject = json::from_file(fileName);
        json::JsonArray array = jsonObject.GetNamedArray(JsonTags::MonitorsTag);

        for (uint32_t i = 0; i < array.Size(); ++i)
        {
            MonitorConfig monitor;
            json::JsonObject value = array.GetObjectAt(i);
            
            monitor.width = (int)value.GetNamedNumber(JsonTags::WidthTag);
            monitor.height = (int)value.GetNamedNumber(JsonTags::HeightTag);
            monitor.left = (int)value.GetNamedNumber(JsonTags::LeftTag);
            monitor.top = (int)value.GetNamedNumber(JsonTags::TopTag);
            monitor.dpi = (int)value.GetNamedNumber(JsonTags::DpiTag);
            monitor.start = value.GetNamedBoolean(JsonTags::UseAsStartMonitorTag);

            result.emplace_back(std::move(monitor));
        }

        return result;
    }

    std::wstring GenerateMonitorId(int width, int height)
    {
        // Unique identifier format: <parsed-device-id>_<width>_<height>_<virtual-desktop-id>
        static const std::wstring virtualDesktopGuid = L"{A96FAA91-BA36-4E5E-B6EB-8DA3B9C30CD5}";
        static int monitorId = 0;
        return L"Monitor#" + std::to_wstring(++monitorId) + 
               L'_' +
               std::to_wstring(width) +
               L'_' +
               std::to_wstring(height) +
               L'_' +
               virtualDesktopGuid;
    }

    std::wstring GenerateEditorArguments(const std::vector<MonitorConfig>& config)
    {
        /*
        * Divider: /
        * Parts:
        * (1) Process id
        * (2) Span zones across monitors
        * (3) Monitor id where the Editor should be opened
        * (4) Monitors count
        *
        * Data for each monitor:
        * (5) Monitor id
        * (6) DPI
        * (7) monitor left
        * (8) monitor top
        * ...
        */
        std::wstring params;
        const std::wstring divider = L"/";
        params += std::to_wstring(GetCurrentProcessId()) + divider; /* Process id */

        const bool spanZonesAcrossMonitors = false;
        params += std::to_wstring(spanZonesAcrossMonitors) + divider; /* Span zones */

        std::wstring targetMonitor;
        std::wstring monitorsData;
        for (const auto& monitor : config)
        {
            auto monitorId = GenerateMonitorId(monitor.width, monitor.height);

            if (monitor.start && targetMonitor.empty())
            {
                params += monitorId + divider; /* Monitor id where the Editor should be opened */
                targetMonitor = monitorId;
            }

            monitorsData += std::move(monitorId) + divider; /* Monitor id */
            monitorsData += std::to_wstring(monitor.dpi) + divider; /* DPI */

            monitorsData += std::to_wstring(monitor.left) + divider;
            monitorsData += std::to_wstring(monitor.top) + divider;
        }

        params += std::to_wstring(config.size()) + divider; /* Monitors count */
        params += monitorsData;

        return params;
    }
}