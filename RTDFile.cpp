/******************************************************************************
*
*   File:   RTDFile.cpp
*
*   Date:   February 5, 2001
*
*   Description:   This file contains the definition of a simple real-time-data 
*                  server for Excel.
*
*	See:
*		- VARIANTs, SAFEARRAYs, and BSTRs, Oh My!
*			http://www.roblocher.com/whitepapers/oletypes.aspx
*		- How To Pass Arrays Between Visual Basic and C
*			http://support.microsoft.com/kb/207931
*		- Array Manipulation Functions
*			http://msdn.microsoft.com/en-us/library/ms221145.aspx
*		- Obtaining Directory Change Notifications
*			http://msdn.microsoft.com/en-us/library/aa365261(VS.85).aspx
*		- map::rend
*			http://msdn.microsoft.com/en-us/library/scb5173k(VS.80).aspx
*
*   Modifications:
******************************************************************************/
#include "RTDFile.h"
#include "RTDFileDLL.h"

#include <stdio.h>
#include <sys/stat.h>
#include <time.h>

#include "atlbase.h"			// CComBSTR
#include <atlconv.h>			// OLE2CA
//#include "stdafx.h"
#include "Tuple.h"

using namespace std;

#pragma warning(disable:4100)

#define RTDSERVERTRACE(x) OutputDebugString(x)
LONG g_cOb = 0;	//global count of the number of objects created.

// Constructor
RTDFile::RTDFile(IUnknown* pUnkOuter)
{
	RTDSERVERTRACE("RTDFile\n");
	m_refCount = 0;
	m_pTypeInfoInterface = NULL;
	m_dwDataThread = MAXDWORD;

	// Get the TypeInfo for this object
	LoadTypeInfo(&m_pTypeInfoInterface, IID_IRtdServer, 0x0);

    // Set up the aggregation
	if (pUnkOuter != NULL)
		m_pOuterUnknown = pUnkOuter;
	else
		m_pOuterUnknown = reinterpret_cast<IUnknown*>
	(static_cast<INonDelegatingUnknown*>(this));   

	// Increment the object count so the server knows not to unload
	InterlockedIncrement( &g_cOb );
}

// Destructor
RTDFile::~RTDFile()
{
   RTDSERVERTRACE("~RTDFile\n");

   // Clean up the type information
   if (m_pTypeInfoInterface != NULL){
      m_pTypeInfoInterface->Release();
      m_pTypeInfoInterface = NULL;
   }

   // Make sure we kill the data thread
   if (m_dwDataThread != -1){
      PostThreadMessage( m_dwDataThread, WM_COMMAND, WM_SILENTTERMINATE, 0 );
   }

   // Decrement the object count
   InterlockedDecrement( &g_cOb );
}

/******************************************************************************
*   LoadTypeInfo -- Gets the type information of an object's interface from the 
*   type library.  Returns S_OK if successful.
******************************************************************************/
STDMETHODIMP RTDFile::LoadTypeInfo(ITypeInfo** pptinfo, REFCLSID clsid,
										  LCID lcid)
{
   //tracing purposes only
   RTDSERVERTRACE("RTDFile::LoadTypeInfo\n");
   
   HRESULT hr;
   LPTYPELIB ptlib = NULL;
   LPTYPEINFO ptinfo = NULL;
   *pptinfo = NULL;
   
   // First try to load the type info from a registered type library
   hr = LoadRegTypeLib(LIBID_RTDServerLib, 1, 0, lcid, &ptlib);
   if (FAILED(hr)){
      RTDSERVERTRACE("Warning: TypeLib not registered.\n");

      //if the libary is not registered, try loading from a file
      hr = LoadTypeLib(LPCOLESTR_RTDFile_DLL, &ptlib);
      if (FAILED(hr)){

         //can't get the type information
         RTDSERVERTRACE("Warning: TypeLib couldn't be loaded.\n");
         return hr;
      }
   }
   
   // Get type information for interface of the object.
   hr = ptlib->GetTypeInfoOfGuid(clsid, &ptinfo);
   if (FAILED(hr))
   {
      ptlib->Release();
      return hr;
   }
   ptlib->Release();
   *pptinfo = ptinfo;
   return S_OK;
}

/******************************************************************************
*   IUnknown Interfaces -- All COM objects must implement, either 
*  directly or indirectly, the IUnknown interface.
******************************************************************************/

/******************************************************************************
*  QueryInterface -- Determines if this component supports the 
*  requested interface, places a pointer to that interface in ppvObj if it's 
*  available, and returns S_OK.  If not, sets ppvObj to NULL and returns 
*  E_NOINTERFACE.
******************************************************************************/
STDMETHODIMP RTDFile::QueryInterface(REFIID riid, void ** ppvObj)
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::QueryInterface\n");

   // defer to the outer unknown
   return m_pOuterUnknown->QueryInterface( riid, ppvObj );
}

/******************************************************************************
*  AddRef() -- In order to allow an object to delete itself when 
*  it is no longer needed, it is necessary to maintain a count of all 
*  references to this object.  When a new reference is created, this function 
*  increments the count.
******************************************************************************/
STDMETHODIMP_(ULONG) RTDFile::AddRef()
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::AddRef\n");
   
   // defer to the outer unknown
   return m_pOuterUnknown->AddRef();
}

/******************************************************************************
*  Release() -- When a reference to this object is removed, this 
*  function decrements the reference count.  If the reference count is 0, then 
*  this function deletes this object and returns 0.
******************************************************************************/
STDMETHODIMP_(ULONG) RTDFile::Release()
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::Release\n");

   // defer to the outer unknown
   return m_pOuterUnknown->Release();
}

/******************************************************************************
*   INonDelegatingUnknown Interfaces -- All COM objects must implement, either 
*  directly or indirectly, the IUnknown interface.
******************************************************************************/

/******************************************************************************
*  NonDelegatingQueryInterface -- Determines if this component supports the 
*  requested interface, places a pointer to that interface in ppvObj if it's 
*  available, and returns S_OK.  If not, sets ppvObj to NULL and returns 
*  E_NOINTERFACE.
******************************************************************************/
STDMETHODIMP RTDFile::NonDelegatingQueryInterface(REFIID riid, void ** ppvObj)
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::NonDelegatingQueryInterface->");
   
   if (riid == IID_IUnknown){
      RTDSERVERTRACE("IUnknown\n");
      *ppvObj = static_cast<INonDelegatingUnknown*>(this);
   }

   else if (riid == IID_IDispatch){
      RTDSERVERTRACE("IDispatch\n");
      *ppvObj = static_cast<IDispatch*>(this);
   }
   
   else if (riid == IID_IRtdServer){
      RTDSERVERTRACE("IRtdServer\n");
      *ppvObj = static_cast<IRtdServer*>(this);
   }
   
   else{
      static char buffer[80];
      LPOLESTR clsidString = NULL;
      StringFromCLSID( riid, &clsidString );
      sprintf( buffer, "Unsupported Interface -- %S\n", clsidString );
      RTDSERVERTRACE( buffer );
      *ppvObj = NULL;
      return E_NOINTERFACE;
   }
   
   static_cast<IUnknown*>(*ppvObj)->AddRef();
   return S_OK;
}

/******************************************************************************
*  NonDelegatingAddRef() -- In order to allow an object to delete itself when 
*  it is no longer needed, it is necessary to maintain a count of all 
*  references to this object.  When a new reference is created, this function 
*  increments the count.
******************************************************************************/
STDMETHODIMP_(ULONG) RTDFile::NonDelegatingAddRef()
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::NonDelegatingAddRef\n");
   
   return ++m_refCount;
}

/******************************************************************************
*  NonDelegatingRelease() -- When a reference to this object is removed, this 
*  function decrements the reference count.  If the reference count is 0, then 
*  this function deletes this object and returns 0.
******************************************************************************/
STDMETHODIMP_(ULONG) RTDFile::NonDelegatingRelease()
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::NonDelegatingRelease\n");
   char txt[10];
   sprintf(txt, "%d", --m_refCount);
   strcat(txt, "\n");
   RTDSERVERTRACE(txt);
   
   if (m_refCount == 0)
   {
      delete this;
      return 0;
   }
   return m_refCount;
}

/******************************************************************************
*   IDispatch Interface -- This interface allows this class to be used as an
*   automation server, allowing its functions to be called by other COM
*   objects
******************************************************************************/

/******************************************************************************
*   GetTypeInfoCount -- This function determines if the class supports type 
*   information interfaces or not.  It places 1 in iTInfo if the class supports
*   type information and 0 if it doesn't.
******************************************************************************/
STDMETHODIMP RTDFile::GetTypeInfoCount(UINT *iTInfo)
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::GetTypeInfoCount\n");
   
   *iTInfo = 0;
   return S_OK;
}

/******************************************************************************
*   GetTypeInfo -- Returns the type information for the class.  For classes 
*   that don't support type information, this function returns E_NOTIMPL;
******************************************************************************/
STDMETHODIMP RTDFile::GetTypeInfo(UINT iTInfo, LCID lcid, 
										 ITypeInfo **ppTInfo)
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::GetTypeInfo\n");

   return E_NOTIMPL;
}

/******************************************************************************
*   GetIDsOfNames -- Takes an array of strings and returns an array of DISPID's
*   which corespond to the methods or properties indicated.  If the name is not 
*   recognized, returns DISP_E_UNKNOWNNAME.
******************************************************************************/
STDMETHODIMP RTDFile::GetIDsOfNames(REFIID riid,  
										   OLECHAR **rgszNames, 
										   UINT cNames,  LCID lcid,
										   DISPID *rgDispId)
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::GetIDsOfNames -- ");
   
   HRESULT hr = E_FAIL;
   
   // Validate arguments
   if (riid != IID_NULL)
      return E_INVALIDARG;
   
   // this API call gets the DISPID's from the type information
   if (m_pTypeInfoInterface != NULL)
      hr = m_pTypeInfoInterface->GetIDsOfNames(rgszNames, cNames, rgDispId);
   
   // DispGetIDsOfNames may have failed, so pass back its return value.
   return hr;
}

/******************************************************************************
*   Invoke -- Takes a dispid and uses it to call another of this class's 
*   methods.  Returns S_OK if the call was successful.
******************************************************************************/
STDMETHODIMP RTDFile::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
									WORD wFlags, DISPPARAMS* pDispParams,
									VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
									UINT* puArgErr)
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::Invoke\n");

   HRESULT hr = DISP_E_PARAMNOTFOUND;
   
   // Validate arguments
   if ((riid != IID_NULL))
      return E_INVALIDARG;

   hr = m_pTypeInfoInterface->Invoke((IRtdServer*)this, dispIdMember, wFlags, 
      pDispParams, pVarResult, pExcepInfo, puArgErr);     

   return S_OK;
}

/******************************************************************************
*  ServerStart -- The ServerStart method is called immediately after a 
*  real-time data server is instantiated.
*  Parameters: CallbackObject -- interface pointer the RTDFile uses to 
*                                indicate new data is available.
*              pfRes -- set to positive value to indicate success.  0 or 
*                       negative value indicates failure.
*  Returns: S_OK
*           E_POINTER
*           E_FAIL
******************************************************************************/
STDMETHODIMP RTDFile::ServerStart(IRTDUpdateEvent *CallbackObject, long *pfRes)
{
	// tracing purposes only
	RTDSERVERTRACE("RTDFile::ServerStart\n");
	HRESULT hr = S_OK;

	// Check the arguments first
	if ((CallbackObject == NULL) || (pfRes == NULL))
		hr = E_POINTER;

	// if the data thread has already been launched, return an error
	else if (m_dwDataThread != -1){
		hr = E_FAIL;
		*pfRes = -1;
	}

	// Try to launch the data thread
	else{
//AfxMessageBox("RTDFile::ServerStart");

		// Marshal the interface to the new thread
		IStream* pMarshalStream = NULL;
		hr = CoMarshalInterThreadInterfaceInStream( IID_IRTDUpdateEvent,
		 CallbackObject, &pMarshalStream );

		CreateThread( NULL, 0, RTDFileThread, (void*)pMarshalStream, 0, &m_dwDataThread );
		*pfRes = m_dwDataThread;
	}

	return hr;
}

HRESULT ParseRTDArgs(SAFEARRAY **ppsa, string* pFilename, string *pFilenameCellref)
{
	HRESULT hr = S_OK;
	USES_CONVERSION;				// Declare local variable used by the OLE2CA macro

	*pFilename = "";
	*pFilenameCellref = "";

	int		argc = 0;
	string	args[2];
	
	// Verify one-dimensional array
	if (SafeArrayGetDim(*ppsa) != 1) return(1);

	// Collect in args, each of the BSTR variants in ppsa
	VARTYPE vt = VT_UNKNOWN;
	if ((hr = SafeArrayGetVartype(*ppsa, &vt)) == S_OK && vt == VT_VARIANT) {

		// Number of arguments (should be 2 - Filename, Cell)
		long lElements= (*ppsa)->rgsabound[0].cElements;

		for (long i = 0; i < min(2, lElements); i++) {
			VARIANT variant;
			if (SafeArrayGetElement(*ppsa, &i, &variant) == S_OK && variant.vt == VT_BSTR) {
				args[argc++] = OLE2CA(variant.bstrVal);
			}
		}
	}

	*pFilename = args[0];
	*pFilenameCellref = Tuple::Create(args[0], args[1]);

	return hr;
}

/******************************************************************************
*  ConnectData -- Adds new topics from a real-time data server. The ConnectData
*  method is called when a file is opened that contains real-time data 
*  functions or when a user types in a new formula which contains the RTD 
*  function.
*  Parameters: TopicID -- value assigned by Excel to identify the topic
*              Strings -- safe array containing the strings identifying the 
*                         data to be served.
*              GetNewValues -- BOOLEAN indicating whether to retrieve new 
*                              values or not.
*              pvarOut -- initial value of the topic
*  Returns: S_OK
*           E_POINTER
*           E_FAIL
******************************************************************************/
STDMETHODIMP RTDFile::ConnectData(long TopicID,
										 SAFEARRAY **Strings,
										 VARIANT_BOOL *GetNewValues,
										 VARIANT *pvarOut)
{
	// tracing purposes only
	RTDSERVERTRACE("RTDFile::ConnectData\n");
	HRESULT hr = S_OK;

	// Check the arguments first
	if (pvarOut == NULL)
		hr = E_POINTER;

	else{
        // Associate TopicID with Filename/Cell strings
		string	Filename;
		string  FilenameCellref;
		ParseRTDArgs(Strings, &Filename, &FilenameCellref);

		struct stat statbuf = {0};
		if (stat(Filename.c_str(), &statbuf) == 0) {
			m_FilenameStat[Filename] = statbuf;
		}

		m_TopicIDMap[TopicID] = FilenameCellref;

		// Lookup the return value
		VariantInit(pvarOut);
		pvarOut->vt = VT_BSTR;
		string CellValue = m_Data.LookupData(FilenameCellref);
		CComBSTR s(CellValue.c_str());
		pvarOut->bstrVal = SysAllocString(s);
	}

	return hr;
}

/******************************************************************************
*  RefreshData -- This method is called by Microsoft Excel to get new data. 
*  This method call only takes place after being notified by the real-time 
*  data server that there is new data.
*  Parameters: TopicCount -- filled with the count of topics in the safearray
*              parrayOut -- two-dimensional safearray.  First dimension 
*                           contains the list of topic IDs.  Second dimension 
*                           contains the values of those topics.
*  Returns: S_OK
*           E_POINTER
*           E_FAIL
******************************************************************************/
STDMETHODIMP RTDFile::RefreshData(long *TopicCount, SAFEARRAY **parrayOut)
{
	// tracing purposes RTDFile
	RTDSERVERTRACE("RTDFile::RefreshData\n");
	HRESULT hr = S_OK;

	// Check the arguments first
	if ((TopicCount == NULL) || (parrayOut == NULL) || (*parrayOut != NULL)){
		hr = E_POINTER;
		RTDSERVERTRACE("   Bad pointer\n");
	}

	else{
		// Set the TopicCount
		*TopicCount = m_TopicIDMap.size();

		SAFEARRAYBOUND bounds[2];
		static WCHAR valBuffer[80];
		VARIANT value;
		long index[2];

		// Build the safe-array values we want to insert sizing for worst case
		bounds[0].cElements = 2;
		bounds[0].lLbound = 0;
		bounds[1].cElements = *TopicCount;
		bounds[1].lLbound = 0;
		*parrayOut = SafeArrayCreate(VT_VARIANT, 2, bounds);

		// Keep track of the new filename stats
		map<string, struct stat> NewFilenameStats;

		map<int, string>::const_iterator iTopicIDMap;
		int i = 0;
		for (iTopicIDMap = m_TopicIDMap.begin(); iTopicIDMap != m_TopicIDMap.end(); iTopicIDMap++) {
			bool dirty = true;

			// Find the FilenameCellref associated with each TopicID
			long TopicID = iTopicIDMap->first;
			string FilenameCellref = iTopicIDMap->second;

			// Extract the filename from m_TopicIDMap
			string Filename;
			if (!Tuple::Get(FilenameCellref, 0, &Filename))
				continue;

			// Find the filename's stat and see if it matches filename's m_FilenameStat
			struct stat statbuf0 = {0};
			if (stat(Filename.c_str(), &statbuf0) == 0) {
				struct stat statbuf1 = m_FilenameStat[Filename];
				dirty = statbuf0.st_mtime != statbuf1.st_mtime ||
						statbuf0.st_ctime != statbuf1.st_ctime ||
						statbuf0.st_size != statbuf1.st_size;
				if (dirty)
					NewFilenameStats[Filename] = statbuf0;
			}

// The diff above causes only 1 I/O Read and the value does not change in Excel!!!
// What's wrong?
//dirty = true;	
			if (dirty) {
				// 
				index[0] = 0;
				index[1] = i;

				VariantInit(&value);
				value.vt = VT_I4;
				value.lVal = TopicID;
				SafeArrayPutElement( *parrayOut, index, &value);

				index[0] = 1;
				index[1] = i;

				// Lookup the return value
				VariantInit(&value);
				value.vt = VT_BSTR;
				string CellValue = m_Data.LookupData(FilenameCellref);
				CComBSTR s(CellValue.c_str());
				value.bstrVal = SysAllocString(s);

				SafeArrayPutElement( *parrayOut, index, &value);

				::VariantClear(& value);

				i++;
			}
		}

		// Update m_FilenameStat with NewFileNameStats
		map<string, struct stat>::const_iterator iNewFilenameStats;
		for (iNewFilenameStats = NewFilenameStats.begin(); iNewFilenameStats != NewFilenameStats.end(); iNewFilenameStats++) {
			m_FilenameStat[iNewFilenameStats->first] = iNewFilenameStats->second;
		}

		// Resize parrayOut
		if (i < *TopicCount) {
			*TopicCount = i;
			bounds[1].cElements = i;
			SafeArrayRedim(*parrayOut, bounds);
		}
	}

	return hr;
}

/******************************************************************************
*  DisconnectData -- Notifies the RTD server application that a topic is no 
*  longer in use.
*  Parameters: TopicID -- the topic that is no longer in use.
*  Returns:
******************************************************************************/
STDMETHODIMP RTDFile::DisconnectData(long TopicID)
{
	// tracing purposes only
	RTDSERVERTRACE("RTDFile::DisconnectData\n");
	HRESULT hr = S_OK;

	// Search for the topic id and remove it
#if 0
// get m_TopicIDMap's Filename, if it is not used in any of the other topics, erase it
	m_FilenameStat.erase();
#endif
	m_TopicIDMap.erase(TopicID);

	return hr;
}

/******************************************************************************
*  Heartbeat -- Determines if the real-time data server is still active.
*  Parameters: pfRes -- filled with zero or negative number to indicate 
*                       failure; positive number indicates success.
*  Returns: S_OK
*           E_POINTER
*           E_FAIL
******************************************************************************/
STDMETHODIMP RTDFile::Heartbeat(long *pfRes)
{
   // tracing purposes only
   RTDSERVERTRACE("RTDFile::Heartbeat\n");
   HRESULT hr = S_OK;

   // Let's reply with the ID of the data thread
   if (pfRes == NULL)
      hr = E_POINTER;
   else
      *pfRes = m_dwDataThread;

   return hr;
}

/******************************************************************************
*  ServerTerminate -- Terminates the connection to the real-time data server.
*  Parameters: none
*  Returns: S_OK
*           E_FAIL
******************************************************************************/
STDMETHODIMP RTDFile::ServerTerminate(void)
{
	// tracing purposes only
	RTDSERVERTRACE("RTDFile::ServerTerminate\n");
	HRESULT hr = S_OK;

	// Make sure we kill the data thread
	if (m_dwDataThread != -1){
		PostThreadMessage( m_dwDataThread, WM_COMMAND, WM_TERMINATE, 0 );
	}

	m_FilenameStat.clear();
	m_TopicIDMap.clear();

	return hr;
}
