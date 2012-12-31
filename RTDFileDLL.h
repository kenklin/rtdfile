/******************************************************************************
*
*	File:	RTDFileDLL.h
*
*	Date:	February 5, 2001
*
*	Description:	This file contains the class factory and standard DLL 
*					functions for the RTDFile COM object.
*
*	Modifications:
******************************************************************************/
#include "comdef.h"
#include "initguid.h"

#define RTDFile_ProgId			"RTDFile"
#define RTDFile_DLL				"RTDFile.dll"
#define LPCOLESTR_RTDFile_DLL	L"RTDFile.dll"
#define RTDFile_Version			"1.0"

class RTDFileClassFactory : public IClassFactory
{
protected:
	ULONG m_refCount;	//reference count

public:
	RTDFileClassFactory();
	~RTDFileClassFactory();

	/******* IUnknown Methods *******/
	STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	/******* IClassFactory Methods *******/
	STDMETHODIMP CreateInstance(LPUNKNOWN, REFIID, LPVOID *);
	STDMETHODIMP LockServer(BOOL);
};

STDAPI DllRegisterServer(void);
STDAPI DllUnregisterServer(void);
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID FAR * ppvObj);
STDAPI DllCanUnloadNow();