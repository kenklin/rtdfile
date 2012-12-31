/******************************************************************************
*
*	File:	RTDFileDLL.cpp
*
*	Date:	February 5, 2001
*
*	Description:	This file contains the class factory and standard DLL 
*					functions for the RTDFile COM object.
*
*	To install:
*		regsvr32 c:\RTDFile\Debug\rtdfile.dll
*
*	References:
*		Real-Time Data: Frequently Asked Questions
*			http://msdn.microsoft.com/en-us/library/aa140060%28office.10%29.aspx
*		Getting started with setup projects
*			http://www.simple-talk.com/dotnet/visual-studio/getting-started-with-setup-projects/
*
*	Modifications:
******************************************************************************/
#include "RTDFileDLL.h"
#include "RTDFile.h"

#define RTDSERVERCLASSFACTORYTRACE(x) OutputDebugString(x)

#define RTDFile_CLSID		"{8D2EEA35-CBEB-49b1-8F3E-68C8F50F38D8}"
const CLSID CLSID_RTDFile = {0x8D2EEA35,0xCBEB,0x49B1,{0x8F,0x3E,0x68,0xC8,0xF5,0x0F,0x38,0xD8}};

LONG g_cLock = 0;	//global count of the locks on this DLL

RTDFileClassFactory::RTDFileClassFactory()
{
   m_refCount = 0;
}

RTDFileClassFactory::~RTDFileClassFactory()
{}

/******************************************************************************
*   IUnknown Interfaces -- All COM objects must implement, either directly or 
*   indirectly, the IUnknown interface.
******************************************************************************/

/******************************************************************************
*   QueryInterface -- Determines if this component supports the requested 
*   interface, places a pointer to that interface in ppvObj if it's available,
*   and returns S_OK.  If not, sets ppvObj to NULL and returns E_NOINTERFACE.
******************************************************************************/
STDMETHODIMP RTDFileClassFactory::QueryInterface(REFIID riid, void ** ppvObj)
{
	//tracing purposes only
	RTDSERVERCLASSFACTORYTRACE("RTDFileClassFactory::QueryInterface->");

	if (riid == IID_IUnknown){
		RTDSERVERCLASSFACTORYTRACE("IUnknown\n");
		*ppvObj = static_cast<IClassFactory*>(this);
	}
   
	else if (riid == IID_IClassFactory){
		RTDSERVERCLASSFACTORYTRACE("IDispatch\n");
		*ppvObj = static_cast<IClassFactory*>(this);
	}
   
	else{
		RTDSERVERCLASSFACTORYTRACE("Unsupported Interface\n");
		*ppvObj = NULL;
		return E_NOINTERFACE;
	}

	static_cast<IUnknown*>(*ppvObj)->AddRef();
	return S_OK;
}

/******************************************************************************
*   AddRef() -- In order to allow an object to delete itself when it is no 
*   longer needed, it is necessary to maintain a count of all references to 
*   this object.  When a new reference is created, this function increments
*   the count.
******************************************************************************/
STDMETHODIMP_(ULONG) RTDFileClassFactory::AddRef()
{
	//tracing purposes only
	RTDSERVERCLASSFACTORYTRACE("RTDFileClassFactory::AddRef\n");
   
	return ++m_refCount;
}

/******************************************************************************
*   Release() -- When a reference to this object is removed, this function 
*   decrements the reference count.  If the reference count is 0, then this 
*   function deletes this object and returns 0;
******************************************************************************/
STDMETHODIMP_(ULONG) RTDFileClassFactory::Release()
{
	//tracing purposes only
	RTDSERVERCLASSFACTORYTRACE("RTDFileClassFactory::Release\n");
   
	if (--m_refCount == 0)
	{
		delete this;
		return 0;
	}
	return m_refCount;
}


/******* IClassFactory Methods *******/
/******************************************************************************
*	CreateInstance() -- This method attempts to create an instance of RTDFile
*	and returns it to the caller.  It maintains a count of the number of
*	created objects.
******************************************************************************/
STDMETHODIMP RTDFileClassFactory::CreateInstance(LPUNKNOWN pUnkOuter, 
												 REFIID riid, 
												 LPVOID *ppvObj)
{
	//tracing purposes only
	RTDSERVERCLASSFACTORYTRACE("RTDFileClassFactory::CreateInstance\n");

	HRESULT hr;
	RTDFile* pObj;

	*ppvObj = NULL;
	hr = ResultFromScode(E_OUTOFMEMORY);

	//It's illegal to ask for anything but IUnknown when aggregating
	if ((pUnkOuter != NULL) && (riid != IID_IUnknown))
		return E_INVALIDARG;
   
	//Create a new instance of RTDFile
	pObj = new RTDFile(pUnkOuter);
   
	if (pObj == NULL)
		return hr;
   
	//Return the resulting object
	hr = pObj->NonDelegatingQueryInterface(riid, ppvObj);

	if (FAILED(hr))
		delete pObj;
   
	return hr;
}

/******************************************************************************
*	LockServer() -- This method maintains a count of the current locks on this
*	DLL.  The count is used to determine if the DLL can be unloaded, or if
*	clients are still using it.
******************************************************************************/
STDMETHODIMP RTDFileClassFactory::LockServer(BOOL fLock)
{
	//tracing purposes only
	RTDSERVERCLASSFACTORYTRACE("RTDFileClassFactory::LockServer\n");

	if (fLock)
		InterlockedIncrement( &g_cLock );
	else
		InterlockedDecrement( &g_cLock );
	return NOERROR;
}

/******* Exported DLL functions *******/
/******************************************************************************
*  g_RegTable -- This N*3 array contains the keys, value names, and values that
*  are associated with this dll in the registry.
******************************************************************************/
const char *g_RegTable[][3] = {
	//format is {key, value name, value }
	{RTDFile_ProgId,			0, RTDFile_ProgId},
	{RTDFile_ProgId "\\CLSID",	0, RTDFile_CLSID},
   
	{"CLSID\\" RTDFile_CLSID,						0, RTDFile_ProgId},
	{"CLSID\\" RTDFile_CLSID "\\InprocServer32",	0, (const char*)-1},
	{"CLSID\\" RTDFile_CLSID "\\ProgId",			0, RTDFile_ProgId},
	{"CLSID\\" RTDFile_CLSID "\\TypeLib",			0, "{0DD8CA71-1832-406a-BCFF-192089D7109A}"},
   
	//	copied this from Kruglinski with my uuids and names.  
	//Just marks where the typelib is
	{"TypeLib\\{0DD8CA71-1832-406a-BCFF-192089D7109A}",									0, RTDFile_ProgId},
	{"TypeLib\\{0DD8CA71-1832-406a-BCFF-192089D7109A}\\" RTDFile_Version,				0, RTDFile_ProgId},
	{"TypeLib\\{0DD8CA71-1832-406a-BCFF-192089D7109A}\\" RTDFile_Version "\\0",			0, "win32"},
	{"TypeLib\\{0DD8CA71-1832-406a-BCFF-192089D7109A}\\" RTDFile_Version "\\0\\Win32",	0, (const char*)-1},
	{"TypeLib\\{0DD8CA71-1832-406a-BCFF-192089D7109A}\\" RTDFile_Version "\\FLAGS",		0, "0"},
};

/******************************************************************************
*  DLLRegisterServer -- This method is the exported method that is used by
*  COM to self-register this component.  It removes the need for a .reg file.
*  ( Taken from Don Box's _Essential COM_ pg. 110-112)
******************************************************************************/
STDAPI DllRegisterServer(void)
{
	HRESULT hr = S_OK;

	//look up server's file name
	char szFileName[255] = "";
	HMODULE dllModule = GetModuleHandle(RTDFile_DLL);
	GetModuleFileName(dllModule, szFileName, 255);
   
	//the typelib should be in the same directory
	char szTypeLibName[255] = "";
	char* pszFileName = NULL;
	memset( szTypeLibName, '\0', 255);
	lstrcpy( szTypeLibName, szFileName );
	pszFileName = strstr( szTypeLibName, RTDFile_DLL);

	//register entries from the table
	int nEntries = sizeof(g_RegTable)/sizeof(*g_RegTable);
	for (int i = 0; SUCCEEDED(hr) && i < nEntries; i++)
	{
		const char *pszName = g_RegTable[i][0];
		const char *pszValueName = g_RegTable[i][1];
		const char *pszValue = g_RegTable[i][2];
      
		//Map rogue values to module file name
		if (pszValue == (const char*)-1)
			pszValue = szFileName;
      
		//Create the key
		HKEY hkey;
		long err = RegCreateKeyA( HKEY_CLASSES_ROOT, pszName, &hkey);
      
		//Set the value
		if (err == ERROR_SUCCESS){
			err = RegSetValueExA( hkey, pszValueName, 0, REG_SZ, 
								(const BYTE*)pszValue, (strlen(pszValue) + 1));
			RegCloseKey(hkey);
		}
      
		//if cannot add key or value, back out and fail
		if (err != ERROR_SUCCESS){
			DllUnregisterServer();
			hr = SELFREG_E_CLASS;
		}
	}
	return hr;
}

/******************************************************************************
*  DllUnregisterServer -- This method is the exported method that is used by 
*  COM to remove the keys added to the registry by DllRegisterServer.  It
*  is essentially for housekeeping.
*  (Taken from Don Box, _Essential COM_ pg 112)
******************************************************************************/
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = S_OK;

	int nEntries = sizeof(g_RegTable)/sizeof(*g_RegTable);

	for (int i = nEntries - 1; i >= 0; i--){
		const char * pszKeyName = g_RegTable[i][0];

		long err = RegDeleteKeyA(HKEY_CLASSES_ROOT, pszKeyName);
		if (err != ERROR_SUCCESS)
			hr = S_FALSE;
	}
	return hr;
}

/******************************************************************************
*	DllGetClassObject() -- This method is the exported method that clients use
*	to create objects in the DLL.  It uses class factories to generate the
*	desired object and returns it to the caller.  The caller must call Release()
*	on the object when they're through with it.
******************************************************************************/
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID FAR * ppvObj)
{
	//tracing purposes only
	RTDSERVERCLASSFACTORYTRACE("DLLGetClassObject");

	//Make sure the requested class is supported by this server
	if (!IsEqualCLSID(rclsid, CLSID_RTDFile))
		return ResultFromScode(E_FAIL);
	   
	//Make sure the requested interface is supported
	if ((!IsEqualCLSID(riid, IID_IUnknown)) && (!IsEqualCLSID(riid, IID_IClassFactory)))
		return ResultFromScode(E_NOINTERFACE);
	   
	//Create the class factory
	*ppvObj = (LPVOID) new RTDFileClassFactory();
	   
	//error checking
	if (*ppvObj == NULL)
		return ResultFromScode(E_OUTOFMEMORY);
	   
	//Addref the Class Factory
	((LPUNKNOWN)*ppvObj)->AddRef();
	   
	return NOERROR;
}

/******************************************************************************
*	DllCanUnloadNow() -- This method checks to see if it's alright to unload 
*	the dll by determining if there are currently any locks on the dll.
******************************************************************************/
STDAPI DllCanUnloadNow()
{
	if ((g_cLock == 0) && (g_cOb == 0))
		return S_OK;
	else
		return S_FALSE;
}