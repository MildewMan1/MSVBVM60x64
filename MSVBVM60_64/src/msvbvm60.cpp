#include "framework.h"
#define CINTERFACE

#include <unknwn.h>
#include <ocidl.h>
#include <string>


//BOOL IsValidInterface(IUnknown* pUnk)
//{
//	BOOL retval = !(IsBadReadPtr(pUnk, sizeof(pUnk)));
//	//size_t s = sizeof(IUnknown);
//	if (!retval)
//		return retval;
//
//	IUnknownVtbl* pVtbl = pUnk->lpVtbl;
//
//	retval = !(IsBadReadPtr(pVtbl, sizeof(pVtbl)));
//	if (!retval)
//		return retval;
//
//	retval = (!IsBadCodePtr((FARPROC)pVtbl->QueryInterface)
//		&& !IsBadCodePtr((FARPROC)pVtbl->AddRef)
//		&& !IsBadCodePtr((FARPROC)pVtbl->Release));
//
//	return retval;
//}

static ULONG __stdcall SafeAddRef(IUnknown* pUnk)
{
	__try
	{
		if (pUnk != nullptr && pUnk->lpVtbl != nullptr && pUnk->lpVtbl->AddRef != nullptr)
		{
			return pUnk->lpVtbl->AddRef(pUnk);
		}
	}
	__except (GetExceptionCode() == EXCEPTION_ACCESS_VIOLATION ?
		EXCEPTION_EXECUTE_HANDLER : EXCEPTION_CONTINUE_SEARCH)
	{
		// A dereference failed because the pointer or vtable was invalid garbage.
	}

	return 0;
}

static ULONG __stdcall SafeRelease(IUnknown* pUnk)
{
	__try
	{
		if (pUnk != nullptr && pUnk->lpVtbl != nullptr && pUnk->lpVtbl->Release != nullptr)
		{
			return pUnk->lpVtbl->Release(pUnk);
		}
	}
	__except (GetExceptionCode() == EXCEPTION_ACCESS_VIOLATION ?
		EXCEPTION_EXECUTE_HANDLER : EXCEPTION_CONTINUE_SEARCH)
	{
		// A dereference failed because the pointer or vtable was invalid garbage.
	}

	return 0;
}

VBA_FUNC(IUnknown*) vbaObjSetByAddress(ULONG_PTR ptr)
{
	/*
	2023-Aug-03 
	Per Microsoft's MSDN (https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-addref):

	"Call this method [AddRef] for every new copy of an interface pointer that you make. 
	For example, if you return a copy of a pointer from a method, then you must 
	call AddRef on that pointer."
	
	*/
	IUnknown* pUnk = reinterpret_cast<IUnknown*>(ptr);

	if (SafeAddRef(pUnk) != 0)
		return pUnk;

	return nullptr;
}

VBA_FUNC(long) vbaGetObjRefCount(ULONG_PTR ptr)
{
	IUnknown* pUnk = reinterpret_cast<IUnknown*>(ptr);
	if (SafeAddRef(pUnk) != 0)
		return static_cast<long>(SafeRelease(pUnk));

	return 0L;
}

