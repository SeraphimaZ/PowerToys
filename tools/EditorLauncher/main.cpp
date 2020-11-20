#include "pch.h"

#include <iostream>
#include <filesystem>

#include "CurrentMonitorConfigurationUtils.h"
#include "CustomMonitorConfiguration.h"

const wchar_t FZEditorExecutablePath[] = L"FancyZonesEditor.exe";

int main(int argc, char* argv[])
{
    std::cout << std::endl;
    std::cout << "---------------------------" << std::endl;
    std::cout << "By default, launcher runs Editor with current monitors configuration." << std::endl << std::endl;
    std::cout << "--example or -e will generate an example input file which you can use as a template for custom monitor configurations." << std::endl;
    std::cout << "--custom or -c will run Editor with custom configuration read from file." << std::endl << std::endl;
    std::cout << "---------------------------" << std::endl << std::endl;

    if (!std::filesystem::exists(FZEditorExecutablePath))
    {
        std::cout << "Error: FancyZonesEditor.exe doesn't exist." << std::endl;
        return -1;
    }

    std::wstring params;
    if (argc == 1)
    {
        params = CurrentMonitorConfiguration::GenerateEditorArguments(); 
    }
    else
    {
        std::string option = std::string(argv[1]);
        if (option == "-e" || option == "--example")
        {
            std::filesystem::path path;
            path.append("example.json");
            CustomMonitorConfiguration::GenerateExample(path.c_str());

            if (!std::filesystem::exists(path))
            {
                std::cout << "Error: example file wasn't created." << std::endl;
                return -1;
            }

            std::cout << "Example is in the " << path << std::endl;
            return 0;
        }
        else if (option == "--custom" || option == "-c")
        {
            if (argc <= 2)
            {
                std::cout << "Error: file name is missing." << std::endl;
                return -1;
            }

            std::string fileName(argv[2]);
            if (fileName.empty() || !std::filesystem::exists(fileName))
            {
                std::cout << "Error: incorrect path to the custom configuration file." << std::endl;
                return -1;
            }

            try
            {
                auto configuration = CustomMonitorConfiguration::ReadConfiguration(argv[2]);
                params = CustomMonitorConfiguration::GenerateEditorArguments(configuration);
            }
            catch (std::exception ex)
            {
                std::cout << "Error: incorrect input file." << std::endl;
                std::cout << ex.what() << std::endl;
                return -1;
            }
        }
    }

    SHELLEXECUTEINFO sei{ sizeof(sei) };
    sei.fMask = { SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI };
    sei.lpFile = FZEditorExecutablePath;
    sei.lpParameters = params.c_str();
    sei.nShow = SW_SHOWDEFAULT;
    ShellExecuteEx(&sei);

    return 0;
}
