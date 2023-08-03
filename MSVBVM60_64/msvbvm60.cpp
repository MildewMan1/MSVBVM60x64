#include "framework.h"
#include "unknwn.h"

#define ExtC extern "C" __declspec(dllexport) 
#define VBA_CALL __stdcall
#define VBA_FUNC(_Type) ExtC _Type VBA_CALL

VBA_FUNC(IUnknown*) vbaObjSetByAddress(IUnknown* pUnknown)
{
	/*
	2023-Aug-03 
	Per Microsoft's MSDN (https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-addref):

	"Call this method [AddRef] for every new copy of an interface pointer that you make. 
	For example, if you return a copy of a pointer from a method, then you must 
	call AddRef on that pointer."
	
	*/

	//check to make sure pUnknown isn't a nullptr before calling AddRef.
	if (pUnknown != nullptr)	
		pUnknown->AddRef();

	return pUnknown;
}