/******************************************************************************
*
*   File:   RTDFile.h
*
*   Date:   February 5, 2001
*
*   Description:   This file contains the declaration of a simple real-time-data 
*                  server for Excel.
*
*   Modifications:
******************************************************************************/
#include "comdef.h"
#include "RTDFileThread.h"
#include "RTDFileData.h"
#include "sys/stat.h"

#include <string>
#include <map>

extern LONG g_cOb;	//global count of the number of objects created.

struct INonDelegatingUnknown
{
	/***** INonDelegatingUnknown Methods *****/
	virtual STDMETHODIMP NonDelegatingQueryInterface(REFIID riid, void ** ppvObj) = 0;      
	virtual STDMETHODIMP_(ULONG) NonDelegatingAddRef() = 0;
	virtual STDMETHODIMP_(ULONG) NonDelegatingRelease() = 0;
};      

class RTDFile : public INonDelegatingUnknown, public IRtdServer
{
private:
	int							m_refCount;
	IUnknown*					m_pOuterUnknown;
	ITypeInfo*					m_pTypeInfoInterface;
	DWORD						m_dwDataThread;

	std::map<int, std::string>			m_TopicIDMap;
	std::map<std::string, struct stat>	m_FilenameStat;
   
public:
	// Constructor
	RTDFile(IUnknown* pUnkOuter);
	// Destructor
	~RTDFile();
   
	STDMETHODIMP LoadTypeInfo(ITypeInfo** pptinfo, REFCLSID clsid, LCID lcid);

	/***** IUnknown Methods *****/
	STDMETHODIMP QueryInterface(REFIID riid, void ** ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	/***** INonDelegatingUnknown Methods *****/
	STDMETHODIMP NonDelegatingQueryInterface(REFIID riid, void ** ppvObj);      
	STDMETHODIMP_(ULONG) NonDelegatingAddRef();
	STDMETHODIMP_(ULONG) NonDelegatingRelease();

	/***** IDispatch Methods *****/
	STDMETHODIMP GetTypeInfoCount(UINT *iTInfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, 
		ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames,  LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, 
		WORD wFlags, DISPPARAMS* pDispParams,
		VARIANT* pVarResult, EXCEPINFO* pExcepInfo,
		UINT* puArgErr);

	/***** IRTD Methods *****/
	STDMETHODIMP ServerStart( 
	  IRTDUpdateEvent *CallbackObject,
	  long *pfRes);

	STDMETHODIMP ConnectData( 
	  long TopicID,
	  SAFEARRAY * *Strings,
	  VARIANT_BOOL *GetNewValues,
	  VARIANT *pvarOut);

	STDMETHODIMP RefreshData( 
	  long *TopicCount,
	  SAFEARRAY * *parrayOut);

	STDMETHODIMP DisconnectData( 
	  long TopicID);

	STDMETHODIMP Heartbeat( 
	  long *pfRes);

	STDMETHODIMP ServerTerminate( void);

private:
	/* RTDFile interface */
	RTDFileData	m_Data;
};
