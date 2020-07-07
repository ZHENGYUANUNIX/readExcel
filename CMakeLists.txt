cmake_minimum_required(VERSION 3.16)
project(readExcel)

set(CMAKE_CXX_STANDARD 14)
link_directories(${PROJECT_SOURCE_DIR})
include_directories(include)
add_executable(${PROJECT_NAME} src/main.cpp)
target_sources(${PROJECT_NAME} PRIVATE src/msgConvert.cpp)
target_link_libraries(${PROJECT_NAME} xlnt)