cmake_minimum_required(VERSION 3.16)
project(wmi_example VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

if(WIN32)
    add_definitions(-DWIN32_LEAN_AND_MEAN)
    add_definitions(-DNOMINMAX)
endif()

include_directories(${CMAKE_SOURCE_DIR}/include)

find_package(Threads REQUIRED)

add_executable(wmi_example
    src/main.cpp
    src/memory.cpp
    src/storage.cpp
)

target_link_libraries(wmi_example
    PRIVATE
    wbemuuid
    ole32
    oleaut32
    Threads::Threads
)

set_target_properties(wmi_example PROPERTIES
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED ON
    WIN32_EXECUTABLE FALSE
)

if(MSVC)
    target_compile_options(wmi_example PRIVATE /W4)
    target_compile_definitions(wmi_example PRIVATE _CRT_SECURE_NO_WARNINGS)
else()
    target_compile_options(wmi_example PRIVATE -Wall -Wextra -Wpedantic)
endif()

if(CMAKE_BUILD_TYPE STREQUAL "Debug")
    target_compile_definitions(wmi_example PRIVATE DEBUG)
else()
    target_compile_definitions(wmi_example PRIVATE NDEBUG)
endif()