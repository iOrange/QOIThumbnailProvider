cmake_minimum_required(VERSION 3.16)

project(QOIThumbnailProvider LANGUAGES CXX)
set(CMAKE_SYSTEM_NAME Windows)
set(CMAKE_CXX_STANDARD 17)

if (MSVC)
    set(CMAKE_STATIC_LINKER_FLAGS "${CMAKE_MODULE_LINKER_FLAGS} /NODEFAULTLIB:LIBCMT")
    set(CMAKE_MODULE_LINKER_FLAGS "${CMAKE_MODULE_LINKER_FLAGS} /DEF:${QOIThumbnailProvider_SOURCE_DIR}/QOIThumbnailProvider.def /NODEFAULTLIB:LIBCMT")
endif()

set(SOURCE_FILES
    Dll.cpp
    Dll.rc
    QOIThumbnailProvider.cpp
)

add_definitions(-DUNICODE)
add_definitions(-D_UNICODE)

add_library(QOIThumbnailProvider MODULE ${SOURCE_FILES})

target_link_libraries(QOIThumbnailProvider PRIVATE shlwapi.lib)
