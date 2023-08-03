#include "pch.h"
#include "unknwn.h"

#include <string>

#define ExtC extern "C" __declspec(dllexport) 
#define VBA_CALL __stdcall
#define VBA_FUNC(type) ExtC type VBA_CALL

#define MY_DEBUG 0

//IUnknown* __stdcall VBAObjSetByAddress(IUnknown* pUnknown, unsigned long* pRefCount)
//{
//
//	if (pUnknown)
//	{
//		unsigned long refCount = pUnknown->AddRef();
//		if (pRefCount != nullptr)
//			*pRefCount = refCount;
//	}
//
//	return pUnknown;
//}

#if MY_DEBUG

unsigned long lastCount{ 0 };
void* pLast{ nullptr };	
#endif

VBA_FUNC(unsigned long) vbaGetObjRefCount(IUnknown* pUnknown)
{
	
	unsigned long retval{ 0 };

	if (pUnknown != nullptr)
	{
		pUnknown->AddRef();
		retval = pUnknown->Release();
	}

#if MY_DEBUG
	static std::wstring msg{};
	msg.reserve(100);

	if (retval > 0 && (IUnknown*)pLast == pUnknown)
	{
		msg.clear();

		msg = L"Last reference count = ";
		msg += std::to_wstring(lastCount);
		msg += L"\n";
		msg += L"Current reference count = ";
		msg += std::to_wstring(retval);
		msg += L"\n";

		OutputDebugString(msg.c_str());
	}
#endif
	return retval;
}

inline IUnknown* __stdcall vbaObjSetByAddressTest(IUnknown* pUnknown, unsigned long* pRefCount = nullptr)
{
	if (pUnknown != nullptr)
	{
		unsigned long refCount = pUnknown->AddRef();

		if (pRefCount != nullptr)
			*pRefCount = refCount;
	}
	return pUnknown;
}

VBA_FUNC(IUnknown*) vbaObjSetByAddress(IUnknown* pUnknown)
{
	
#if MY_DEBUG
	IUnknown* retval = vbaObjSetByAddressTest(pUnknown, &lastCount);
	pLast = (void*)retval;
	return retval;
#else
	return vbaObjSetByAddressTest(pUnknown);
#endif
	
}