#pragma once
#ifndef __FRAMEWORK_H__
#define __FRAMEWORK_H__

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#define ExtC extern "C" __declspec(dllexport) 
#define VBA_CALL __stdcall
#define VBA_FUNC(_Type) ExtC _Type VBA_CALL

#endif //__FRAMEWORK_H__
