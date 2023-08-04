// This file was originally by Matt Pietrek 1999
// in Microsoft Systems Journal, March 1999
//
// It enumerates COM interface methods
//
// See http://www.microsoft.com/msj/0399/comtype/comtype.htm 
// for the full article
//
// I added some code (error handling, etc.)
// and deleted some code not used. Dieter Spaar
#include "stdafx.h"
#include <windows.h>
#include <ole2.h>
#include <ocidl.h>
#include <imagehlp.h>
#include <tchar.h>
#include <sstream>

#include "CoClassSyms.h"
#include "cJSON.h"

#ifndef UNICODE
#error "This file must be compiled as unicode"
#endif
static DWORD pb = 0;
// internal declarations

void ShowHresultError(const char *szErr, HRESULT hr);
void EnumTypeLib(LPTYPELIB pITypeLib);
void ProcessTypeInfo(LPTYPEINFO pITypeInfo);
void ProcessReferencedTypeInfo(LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr,
	HREFTYPE hRefType);
void EnumTypeInfoMembers(LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr,
	LPUNKNOWN lpUnknown);
void GetTypeInfoName(LPTYPEINFO pITypeInfo, LPTSTR pszName, LPTSTR pszComment = NULL,
	MEMBERID memid = MEMBERID_NIL);

extern cJSON* printJson;
//============================================================================
// display error messages if a COM function failed
//============================================================================
void ShowHresultError(const char *szErr, HRESULT hr)
{
	LPVOID lpMsgBuf = 0;

	DWORD dw = FormatMessage(
		FORMAT_MESSAGE_ALLOCATE_BUFFER |
		FORMAT_MESSAGE_FROM_SYSTEM |
		FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,
		hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
		(LPTSTR)&lpMsgBuf,
		0,
		NULL
		);
	char szBuf[1024 * 8];
	if (dw == 0 || lpMsgBuf == 0)
	{
		if (hr == CLASS_E_NOTLICENSED) // seems that this message is not available on some systems
			wsprintfA(szBuf, "%s: (0x%X) %ls", szErr, hr, _T("Class is not licensed for use"));
		else
			wsprintfA(szBuf, "%s: (0x%X) %ls", szErr, hr, _T("Error message not found"));
	}
	else
		wsprintfA(szBuf, "%s: (0x%X) %ls", szErr, hr, (LPCTSTR)lpMsgBuf);

	if (lpMsgBuf)
		LocalFree(lpMsgBuf);

	// remove linefeeds, displayed as garbage in IDA otherwise

	for (;;)
	{
		char *p = strchr(szBuf, '\r');
		if (p)
			*p = ' ';
		else
			break;
	}

	// add some suggestion how the problem can be resolved

	if (hr == CLASS_E_NOTLICENSED)
		strcat(szBuf, "\nThe COM component needs a valid license to be used.");
	else if (hr == REGDB_E_CLASSNOTREG)
		strcat(szBuf, "\nThis may indicate that function addresses cannot be retrieved for this interface, but you can try to register the component with REGSVR32.");

	printf(szBuf);
}

//============================================================================
// Inizialization
//============================================================================
BOOL InitProcessTypeLib()
{
	HRESULT hr = CoInitialize(0);		// Initialize COM subsystem
	if (S_OK != hr && S_FALSE != hr)
	{
		ShowHresultError("CoInitialize() failed", hr);
		return FALSE;
	}

	return TRUE;
}

//============================================================================
// Cleanup
//============================================================================
BOOL ExitProcessTypeLib()
{
	CoUninitialize();

	return TRUE;
}

//============================================================================
// Given a filename for a typelib, attempt to get an ITypeLib for it.  Send
// the resultant ITypeLib instance to EnumTypeLib.
//============================================================================
void ProcessTypeLib(BSTR pszFileName)
{
	
	LPTYPELIB pITypeLib;

	HRESULT hr = LoadTypeLib(pszFileName, &pITypeLib);
	if (S_OK != hr)
	{
		ShowHresultError("LoadTypeLib() failed", hr);
		return;
	}

	EnumTypeLib(pITypeLib);

	pITypeLib->Release();
	
}

//============================================================================
// Enumerate through all the ITypeInfo instances in an ITypeLib.  Pass each
// instance to ProcessTypeInfo.
//============================================================================
void EnumTypeLib(LPTYPELIB pITypeLib)
{
	UINT tiCount = pITypeLib->GetTypeInfoCount();

	for (UINT i = 0; i < tiCount; i++)
	{
		LPTYPEINFO pITypeInfo;

		HRESULT hr = pITypeLib->GetTypeInfo(i, &pITypeInfo);

		if (S_OK == hr)
		{
			ProcessTypeInfo(pITypeInfo);

			pITypeInfo->Release();
		}
		else
			ShowHresultError("GetTypeInfo() failed", hr);
	}
}

//============================================================================
// Top level handling code for a single ITypeInfo extracted from a typelib
//============================================================================
void ProcessTypeInfo(LPTYPEINFO pITypeInfo)
{
	HRESULT hr;

	LPTYPEATTR pTypeAttr;
	hr = pITypeInfo->GetTypeAttr(&pTypeAttr);
	if (S_OK != hr)
	{
		ShowHresultError("GetTypeAttr() failed", hr);
		return;
	}
	
	if (TKIND_COCLASS == pTypeAttr->typekind)
	{
		for (unsigned short i = 0; i < pTypeAttr->cImplTypes; i++)
		{
			HREFTYPE hRefType;
			//DecodePointer(IMR_RECONVERTSTRING);
			hr = pITypeInfo->GetRefTypeOfImplType(i, &hRefType);

			if (S_OK == hr)
				ProcessReferencedTypeInfo(pITypeInfo, pTypeAttr, hRefType);
			else
				ShowHresultError("GetRefTypeOfImplType() failed", hr);
		}
	}

	pITypeInfo->ReleaseTypeAttr(pTypeAttr);
}

//============================================================================
// Given a TKIND_COCLASS ITypeInfo, get the ITypeInfo that describes the
// referenced (HREFTYPE) TKIND_DISPATCH or TKIND_INTERFACE.  Pass that
// ITypeInfo to EnumTypeInfoMembers. 3,362.5	 2550/2=1275
//============================================================================
void ProcessReferencedTypeInfo(LPTYPEINFO pITypeInfo_CoClass,
	LPTYPEATTR pTypeAttr,
	HREFTYPE hRefType)
{
	LPTYPEINFO pIRefTypeInfo;	

	HRESULT hr = pITypeInfo_CoClass->GetRefTypeInfo(hRefType, &pIRefTypeInfo);
	if (S_OK != hr)
	{
		ShowHresultError("GetRefTypeInfo() failed", hr);
		return;
	}
	LPTYPEATTR pRefTypeAttr = 0;
	hr = pIRefTypeInfo->GetTypeAttr(&pRefTypeAttr);
	if (S_OK != hr)
	{
		pIRefTypeInfo->Release();
		ShowHresultError("GetTypeAttr() failed", hr);
		return;
	}
	if (pRefTypeAttr == 0)
	{
		pIRefTypeInfo->ReleaseTypeAttr(pRefTypeAttr);
		pIRefTypeInfo->Release();
		ShowHresultError("pRefTypeAttr == 0", S_OK);
		return;
	}

	LPUNKNOWN pIUnknown = 0;

	// may return CLASS_E_NOTLICENSED

	hr = CoCreateInstance(pTypeAttr->guid,
		0,					// pUnkOuter
		CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER,
		pRefTypeAttr->guid,
		(LPVOID *)&pIUnknown);

	if ((S_OK == hr) && pIUnknown)
	{
		EnumTypeInfoMembers(pIRefTypeInfo, pRefTypeAttr, pIUnknown);

		pIUnknown->Release();
	}
	if (S_OK != hr && E_NOINTERFACE != hr)
	{
		// get interface name for error message

		TCHAR pszInterfaceName[256];
		GetTypeInfoName(pIRefTypeInfo, pszInterfaceName);
		char szBuf[256];
		wsprintfA(szBuf, "CoCreateInstance() for %ls failed", pszInterfaceName);

		ShowHresultError(szBuf, hr);
	}

	pIRefTypeInfo->ReleaseTypeAttr(pRefTypeAttr);
	pIRefTypeInfo->Release();
}

BOOL CheckInterface(LPTYPEINFO pITypeInfo, LPTYPEATTR pTypeAttr, const char *szInterfaceName)
{
	// IDispatch interface must be a dual interface, otherwise the method address is
	// the same for all methodes of that interface0

	if (pTypeAttr->typekind == TKIND_DISPATCH && (pTypeAttr->wTypeFlags & TYPEFLAG_FDUAL) == 0)
	{
		printf("IDispatch interface %s is not dual.\n", szInterfaceName);
		return FALSE;
	}

	return TRUE;
}
//dword to hex
std::string DwordtoHex(DWORD pFunction){
	std::ostringstream os;
	os << std::hex << pFunction;
	return os.str();
}

void ProcessFunction(cJSON* interfaceJson,FUNCDESC *pFuncDesc, LPTYPEINFO pITypeInfo, DWORD pFunction, const char * pszInterfaceNameAnsi, const char *pszMungedName, const char *pszCommentAnsi)
{	
	
	LPUNKNOWN lpUnknown = 0;
//	DWORD pFunction;
	//PBYTE pVTable = (PBYTE)*(PDWORD)(lpUnknown);

	
	cJSON* funcJson = cJSON_CreateObject();
	cJSON_AddStringToObject(funcJson, "name", pszMungedName);	
	//cJSON_AddStringToObject(funcJson, "address", DwordtoHex(pFunction).c_str());	
	//cJSON_AddStringToO(funcJson,"typeinfo", pITypeInfo);
	cJSON_AddStringToObject(funcJson,"address",DwordtoHex(pFunction-pb).c_str());
	cJSON_AddItemToArray(interfaceJson,  funcJson);
}

//============================================================================
// Enumerate through each member of an ITypeInfo.  Send the method name and
// address to the CoClassSymsAddSymbol function.
//=============================================================================
void EnumTypeInfoMembers(LPTYPEINFO pITypeInfo,	// The ITypeInfo to enum.
	LPTYPEATTR pTypeAttr,	// The associated TYPEATTR.
	LPUNKNOWN lpUnknown		// From CoCreateInstance.
	)
{
	// Make a pointer to the vtable.	
	PBYTE pVTable = (PBYTE)*(PDWORD)(lpUnknown);

	if (0 == pTypeAttr->cFuncs)	// Make sure at least one method!
		return;

	// Get the name of the ITypeInfo, to use as the interface name in the
	// symbol names we'll be constructing.
	TCHAR pszInterfaceName[256];
	GetTypeInfoName(pITypeInfo, pszInterfaceName);
	char pszInterfaceNameAnsi[256];
	wsprintfA(pszInterfaceNameAnsi, "%ls", pszInterfaceName);
	cJSON* interfaceJson = cJSON_CreateArray();
	if (!cJSON_GetObjectItem(printJson, "modulebase"))
	{
		
		MEMORY_BASIC_INFORMATION mb = { 0 };
		VirtualQuery(pVTable, &mb, sizeof(mb));
		pb = (DWORD)mb.AllocationBase;
		//DwordtoHex(pFunction).c_str()
		cJSON_AddStringToObject(printJson, "modulebase", DwordtoHex(pb).c_str());
	}
	cJSON_AddItemToObject(printJson, pszInterfaceNameAnsi, interfaceJson);
	// check if interface should be proccessed
	
	if (!CheckInterface(pITypeInfo, pTypeAttr, pszInterfaceNameAnsi))
		return;

	// Enumerate through each method, obtain it's name, address, and ship the
	// info off to CoClassSymsAddSymbol()
	for (unsigned i = 0; i < pTypeAttr->cFuncs; i++)
	{
		FUNCDESC *pFuncDesc;

		HRESULT hr = pITypeInfo->GetFuncDesc(i, &pFuncDesc);
		if (S_OK != hr)
		{
			ShowHresultError("GetFuncDesc() failed", hr);
			continue;
		}

		TCHAR pszMemberName[256];
		TCHAR pszComment[256];
		GetTypeInfoName(pITypeInfo, pszMemberName, pszComment, pFuncDesc->memid);

		// Index into the vtable to retrieve the method's virtual address
		DWORD pFunction = *(PDWORD)(pVTable + pFuncDesc->oVft);

		// Created the basic form of the symbol name in interface::method
		// form using ANSI characters
		char pszMungedNameAnsi[512];
		wsprintfA(pszMungedNameAnsi, "%ls::%ls", pszInterfaceName, pszMemberName);

		char pszCommentAnsi[512];
		wsprintfA(pszCommentAnsi, "%ls", pszComment);

		INVOKEKIND invkind = pFuncDesc->invkind;

		// If it's a property "get" or "put", append a meaningful ending.
		// The "put" and "get" will have identical names, so we want to
		// make them into unique names
		if (INVOKE_PROPERTYGET == invkind)
			strcat(pszMungedNameAnsi, "_get");
		else if (INVOKE_PROPERTYPUT == invkind)
			strcat(pszMungedNameAnsi, "_put");
		else if (INVOKE_PROPERTYPUTREF == invkind)
			strcat(pszMungedNameAnsi, "_putref");

		// process function

		ProcessFunction(interfaceJson,pFuncDesc, pITypeInfo, pFunction, pszInterfaceNameAnsi, pszMungedNameAnsi, pszCommentAnsi);

		pITypeInfo->ReleaseFuncDesc(pFuncDesc);
	}
}

//============================================================================
// Given an ITypeInfo instance, retrieve the name.
//=============================================================================
void GetTypeInfoName(LPTYPEINFO pITypeInfo, LPTSTR pszName, LPTSTR pszComment, MEMBERID memid)
{
	BSTR pszTypeInfoName;
	BSTR pszDocString;
	HRESULT hr;

	if (pszComment)
		pszComment[0] = _T('\0');

	hr = pITypeInfo->GetDocumentation(memid, &pszTypeInfoName, &pszDocString, 0, 0);

	if (S_OK != hr)
	{
		lstrcpy(pszName, _T("<unknown>"));
		return;
	}

	// Make a copy so that we can free the BSTR	allocated by ::GetDocumentation
	if (pszTypeInfoName){
		lstrcpyW(pszName, pszTypeInfoName);
	}
	if (pszComment&&pszDocString)
		lstrcpyW(pszComment, pszDocString);

	// Free the BSTR allocated by ::GetDocumentation
	if (pszTypeInfoName){
		SysFreeString(pszTypeInfoName);
	}
	if (pszDocString)
	{
		SysFreeString(pszDocString);
	}
}
