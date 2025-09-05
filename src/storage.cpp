#include <iomanip>
#include <iostream>
#include <memory>
#include <wmi/wmi.hxx>

void QueryStorageInfo(std::shared_ptr<const wmi::Interface> wmi_interface) {
    try {
        std::cout << "Querying logical disk information..." << std::endl;
        auto disk_result = wmi_interface->ExecuteQuery(
            L"SELECT DeviceID, Size, FreeSpace, FileSystem, DriveType FROM Win32_LogicalDisk");

        for (const auto& disk_obj : disk_result) {
            auto device_id = disk_obj.GetProperty<std::string>(L"DeviceID");
            auto size = disk_obj.GetProperty<std::string>(L"Size");
            auto free_space = disk_obj.GetProperty<std::string>(L"FreeSpace");
            auto file_system = disk_obj.GetProperty<std::string>(L"FileSystem");
            auto drive_type = disk_obj.GetProperty<std::string>(L"DriveType");

            if (device_id) {
                std::cout << "  Drive " << *device_id << ":" << std::endl;

                if (size && free_space) {
                    double size_gb = std::stod(*size) / (1024.0 * 1024.0 * 1024.0);
                    double free_gb = std::stod(*free_space) / (1024.0 * 1024.0 * 1024.0);
                    double used_gb = size_gb - free_gb;
                    double usage_percent = (used_gb / size_gb) * 100.0;

                    std::cout << std::fixed << std::setprecision(2);
                    std::cout << "    Total Size: " << size_gb << " GB" << std::endl;
                    std::cout << "    Free Space: " << free_gb << " GB" << std::endl;
                    std::cout << "    Used Space: " << used_gb << " GB" << std::endl;
                    std::cout << "    Usage:      " << usage_percent << "%" << std::endl;
                }

                if (file_system) {
                    std::cout << "    File System: " << *file_system << std::endl;
                }

                if (drive_type) {
                    std::string type_name;
                    int type_code = std::stoi(*drive_type);
                    switch (type_code) {
                        case 0:
                            type_name = "Unknown";
                            break;
                        case 1:
                            type_name = "No Root Directory";
                            break;
                        case 2:
                            type_name = "Removable Disk";
                            break;
                        case 3:
                            type_name = "Local Disk";
                            break;
                        case 4:
                            type_name = "Network Drive";
                            break;
                        case 5:
                            type_name = "Compact Disc";
                            break;
                        case 6:
                            type_name = "RAM Disk";
                            break;
                        default:
                            type_name = "Unknown Type";
                            break;
                    }
                    std::cout << "    Drive Type: " << type_name << std::endl;
                }

                std::cout << std::endl;
            }
        }

        std::cout << std::endl;

        std::cout << "Querying physical disk information..." << std::endl;
        auto physical_disk_result = wmi_interface->ExecuteQuery(
            L"SELECT Model, Size, MediaType, InterfaceType FROM Win32_DiskDrive");

        int disk_count = 0;
        for (const auto& physical_disk_obj : physical_disk_result) {
            disk_count++;

            auto model = physical_disk_obj.GetProperty<std::string>(L"Model");
            auto size = physical_disk_obj.GetProperty<std::string>(L"Size");
            auto media_type = physical_disk_obj.GetProperty<std::string>(L"MediaType");
            auto interface_type = physical_disk_obj.GetProperty<std::string>(L"InterfaceType");

            std::cout << "  Physical Disk " << disk_count << ":" << std::endl;

            if (model) {
                std::cout << "    Model: " << *model << std::endl;
            }

            if (size) {
                double size_gb = std::stod(*size) / (1024.0 * 1024.0 * 1024.0);
                std::cout << "    Size: " << std::fixed << std::setprecision(2) << size_gb << " GB"
                          << std::endl;
            }

            if (media_type) {
                std::cout << "    Media Type: " << *media_type << std::endl;
            }

            if (interface_type) {
                std::cout << "    Interface: " << *interface_type << std::endl;
            }

            std::cout << std::endl;
        }

        if (disk_count == 0) {
            std::cout << "  No physical disks found." << std::endl;
        }

    } catch (const wmi::Exception& e) {
        std::cerr << "Storage query error: " << e.what() << std::endl;
    } catch (const std::exception& e) {
        std::cerr << "Storage query error: " << e.what() << std::endl;
    }
}