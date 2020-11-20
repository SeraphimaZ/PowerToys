#include "CurrentMonitorConfigurationUtils.h"
#include "pch.h"

#include <optional>
#include <vector>

#include <ShellScalingApi.h>

struct Rect
{
    Rect() {}

    Rect(RECT rect) :
        m_rect(rect)
    {
    }

    Rect(RECT rect, UINT dpi) :
        m_rect(rect)
    {
        m_rect.right = m_rect.left + MulDiv(m_rect.right - m_rect.left, dpi, 96);
        m_rect.bottom = m_rect.top + MulDiv(m_rect.bottom - m_rect.top, dpi, 96);
    }

    int x() const { return m_rect.left; }
    int y() const { return m_rect.top; }
    int width() const { return m_rect.right - m_rect.left; }
    int height() const { return m_rect.bottom - m_rect.top; }
    int left() const { return m_rect.left; }
    int top() const { return m_rect.top; }
    int right() const { return m_rect.right; }
    int bottom() const { return m_rect.bottom; }
    int aspectRatio() const { return MulDiv(m_rect.bottom - m_rect.top, 100, m_rect.right - m_rect.left); }

private:
    RECT m_rect{};
};

std::wstring TrimDeviceId(const std::wstring& deviceId)
{
    // We're interested in the unique part between the first and last #'s
    // Example input: \\?\DISPLAY#DELA026#5&10a58c63&0&UID16777488#{e6f07b5f-ee97-4a90-b076-33f57bf4eaa7}
    // Example output: DELA026#5&10a58c63&0&UID16777488
    static const std::wstring defaultDeviceId = L"FallbackDevice";
    if (deviceId.empty())
    {
        return defaultDeviceId;
    }

    size_t start = deviceId.find(L'#');
    size_t end = deviceId.rfind(L'#');
    if (start != std::wstring::npos && end != std::wstring::npos && start != end)
    {
        size_t size = end - (start + 1);
        return deviceId.substr(start + 1, size);
    }
    else
    {
        return defaultDeviceId;
    }
}

std::wstring GenerateUniqueId(HMONITOR monitor, const std::wstring& deviceId, const std::wstring& virtualDesktopId)
{
    MONITORINFOEXW mi;
    mi.cbSize = sizeof(mi);
    if (!virtualDesktopId.empty() && GetMonitorInfo(monitor, &mi))
    {
        Rect const monitorRect(mi.rcMonitor);
        // Unique identifier format: <parsed-device-id>_<width>_<height>_<virtual-desktop-id>
        return TrimDeviceId(deviceId) +
               L'_' +
               std::to_wstring(monitorRect.width()) +
               L'_' +
               std::to_wstring(monitorRect.height()) +
               L'_' +
               virtualDesktopId;
    }
    return {};
}

std::optional<std::wstring> GenerateMonitorId(MONITORINFOEX mi, HMONITOR monitor, const GUID& virtualDesktopId)
{
    DISPLAY_DEVICE displayDevice = { sizeof(displayDevice) };
    PCWSTR deviceId = nullptr;

    bool validMonitor = true;
    if (EnumDisplayDevices(mi.szDevice, 0, &displayDevice, 1))
    {
        if (displayDevice.DeviceID[0] != L'\0')
        {
            deviceId = displayDevice.DeviceID;
        }
    }

    if (!deviceId)
    {
        deviceId = GetSystemMetrics(SM_REMOTESESSION) ?
                       L"\\\\?\\DISPLAY#REMOTEDISPLAY#" :
                       L"\\\\?\\DISPLAY#LOCALDISPLAY#";
    }

    std::wstring guidStr;
    OLECHAR* vdId;
    if (StringFromCLSID(virtualDesktopId, &vdId) == S_OK)
    {
        guidStr = vdId;
    }

    CoTaskMemFree(vdId);
    return GenerateUniqueId(monitor, deviceId, guidStr);
}

template<RECT MONITORINFO::*member>
std::vector<std::pair<HMONITOR, MONITORINFOEX>> GetAllMonitorInfo()
{
    using result_t = std::vector<std::pair<HMONITOR, MONITORINFOEX>>;
    result_t result;

    auto enumMonitors = [](HMONITOR monitor, HDC hdc, LPRECT pRect, LPARAM param) -> BOOL {
        MONITORINFOEX mi;
        mi.cbSize = sizeof(mi);
        result_t& result = *reinterpret_cast<result_t*>(param);
        if (GetMonitorInfo(monitor, &mi))
        {
            result.push_back({ monitor, mi });
        }

        return TRUE;
    };

    EnumDisplayMonitors(NULL, NULL, enumMonitors, reinterpret_cast<LPARAM>(&result));
    return result;
}

std::wstring CurrentMonitorConfiguration::GenerateEditorArguments()
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

    std::vector<std::pair<HMONITOR, MONITORINFOEX>> allMonitors;
    allMonitors = GetAllMonitorInfo<&MONITORINFOEX::rcWork>();
    HMONITOR targetMonitor = MonitorFromWindow(GetForegroundWindow(), MONITOR_DEFAULTTOPRIMARY);

    GUID virtualDesktopId{};

    std::wstring monitorsData;
    for (auto& monitor : allMonitors)
    {
        auto monitorId = GenerateMonitorId(monitor.second, monitor.first, virtualDesktopId);
        if (monitor.first == targetMonitor)
        {
            params += *monitorId + divider; /* Monitor id where the Editor should be opened */
        }

        if (monitorId.has_value())
        {
            monitorsData += std::move(*monitorId) + divider; /* Monitor id */

            UINT dpiX = 0;
            UINT dpiY = 0;
            if (GetDpiForMonitor(monitor.first, MDT_EFFECTIVE_DPI, &dpiX, &dpiY) == S_OK)
            {
                monitorsData += std::to_wstring(dpiX) + divider; /* DPI */
            }

            monitorsData += std::to_wstring(monitor.second.rcMonitor.left) + divider;
            monitorsData += std::to_wstring(monitor.second.rcMonitor.top) + divider;
        }
    }

    params += std::to_wstring(allMonitors.size()) + divider; /* Monitors count */
    params += monitorsData;

    return params;
}
