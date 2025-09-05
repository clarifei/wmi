#include <comdef.h>

#include <chrono>
#include <iostream>
#include <memory>
#include <stdexcept>
#include <wmi/wmi.hxx>

void QueryMemoryInfo(std::shared_ptr<const wmi::Interface> wmi_interface);
void QueryStorageInfo(std::shared_ptr<const wmi::Interface> wmi_interface);

int main() {
    try {
        std::cout << "WMI++ Library Example\n" << std::endl;

        const auto start_time = std::chrono::high_resolution_clock::now();

        wmi::COMInitializer com_init(COINIT_MULTITHREADED);

        if (com_init.IsInitialized()) {
            std::cout << "COM library initialized successfully by COMInitializer." << std::endl;
        } else {
            std::cout << "COM library was already initialized (using existing initialization)."
                      << std::endl;
        }
        std::cout << std::endl;

        auto wmi_interface = wmi::Interface::Create("cimv2");
        if (!wmi_interface) {
            std::cerr << "Failed to create WMI interface" << std::endl;
            return 1;
        }

        std::cout << "WMI interface created successfully!\n" << std::endl;

        std::cout << "Memory Information" << std::endl;
        QueryMemoryInfo(wmi_interface);
        std::cout << std::endl;

        std::cout << "Storage Information" << std::endl;
        QueryStorageInfo(wmi_interface);
        std::cout << std::endl;

        const auto end_time = std::chrono::high_resolution_clock::now();
        const auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);

        std::cout << "All queries completed successfully! (" << duration.count() << " ms)" << std::endl;

    } catch (const wmi::Exception& e) {
        std::cerr << "WMI Error: " << e.what() << std::endl;
        return 1;
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    } catch (...) {
        std::cerr << "Unknown error occurred" << std::endl;
        return 1;
    }

    return 0;
}