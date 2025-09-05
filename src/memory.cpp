#include <iomanip>
#include <iostream>
#include <memory>
#include <wmi/wmi.hxx>

void QueryMemoryInfo(std::shared_ptr<const wmi::Interface> wmi_interface) {
    try {
        std::cout << "Querying operating system memory information..." << std::endl;
        auto os_result = wmi_interface->ExecuteQuery(
            L"SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem");

        for (const auto& os_obj : os_result) {
            auto total_memory = os_obj.GetProperty<std::string>(L"TotalVisibleMemorySize");
            auto free_memory = os_obj.GetProperty<std::string>(L"FreePhysicalMemory");

            if (total_memory && free_memory) {
                double total_mb = std::stod(*total_memory) / 1024.0;
                double free_mb = std::stod(*free_memory) / 1024.0;
                double used_mb = total_mb - free_mb;
                double usage_percent = (used_mb / total_mb) * 100.0;

                std::cout << std::fixed << std::setprecision(2);
                std::cout << "  Total Physical Memory: " << total_mb << " MB" << std::endl;
                std::cout << "  Free Physical Memory:  " << free_mb << " MB" << std::endl;
                std::cout << "  Used Physical Memory:  " << used_mb << " MB" << std::endl;
                std::cout << "  Memory Usage:          " << usage_percent << "%" << std::endl;
            }
        }

        std::cout << std::endl;

        std::cout << "Querying physical memory modules..." << std::endl;
        auto memory_result = wmi_interface->ExecuteQuery(
            L"SELECT Capacity, Speed, Manufacturer, PartNumber FROM Win32_PhysicalMemory");

        int module_count = 0;
        for (const auto& memory_obj : memory_result) {
            module_count++;

            auto capacity = memory_obj.GetProperty<std::string>(L"Capacity");
            auto speed = memory_obj.GetProperty<std::string>(L"Speed");
            auto manufacturer = memory_obj.GetProperty<std::string>(L"Manufacturer");
            auto part_number = memory_obj.GetProperty<std::string>(L"PartNumber");

            std::cout << "  Module " << module_count << ":" << std::endl;

            if (capacity) {
                double capacity_gb = std::stod(*capacity) / (1024.0 * 1024.0 * 1024.0);
                std::cout << "    Capacity: " << std::fixed << std::setprecision(1) << capacity_gb
                          << " GB" << std::endl;
            }

            if (speed) {
                std::cout << "    Speed: " << *speed << " MHz" << std::endl;
            }

            if (manufacturer) {
                std::cout << "    Manufacturer: " << *manufacturer << std::endl;
            }

            if (part_number) {
                std::cout << "    Part Number: " << *part_number << std::endl;
            }

            std::cout << std::endl;
        }

        if (module_count == 0) {
            std::cout << "  No physical memory modules found." << std::endl;
        }

    } catch (const wmi::Exception& e) {
        std::cerr << "Memory query error: " << e.what() << std::endl;
    } catch (const std::exception& e) {
        std::cerr << "Memory query error: " << e.what() << std::endl;
    }
}