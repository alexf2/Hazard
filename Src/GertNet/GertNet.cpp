// GertNet.cpp : Implementation of DLL Exports.


// Note: Proxy/Stub Information
//      To merge the proxy/stub code into the object DLL, add the file 
//      dlldatax.c to the project.  Make sure precompiled headers 
//      are turned off for this file, and add _MERGE_PROXYSTUB to the 
//      defines for the project.  
//
//      If you are not running WinNT4.0 or Win95 with DCOM, then you
//      need to remove the following define from dlldatax.c
//      #define _WIN32_WINNT 0x0400
//
//      Further, if you are running MIDL without /Oicf switch, you also 
//      need to remove the following define from dlldatax.c.
//      #define USE_STUBLESS_PROXY
//
//      Modify the custom build rule for GertNet.idl by adding the following 
//      files to the Outputs.
//          GertNet_p.c
//          dlldata.c
//      To build a separate proxy/stub DLL, 
//      run nmake -f GertNetps.mk in the project directory.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include "GertNet.h"
#include "dlldatax.h"

#include "GertNet_i.c"
#include "LingvoEnum.h"
#include "CollLingvo.h"
#include "CompaundFiles.h"
#include "Factor.h"
#include "CollFac.h"
#include "MGertNet.h"
#include "SafetyPrecaution.h"
#include "FChange.h"
#include "ICollFChange.h"
#include "CollSF.h"
#include "CatRegistrar.h"

#ifdef _MERGE_PROXYSTUB
extern "C" HINSTANCE hProxyDll;
#endif

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_LingvoEnum, CLingvoEnum)
OBJECT_ENTRY(CLSID_CollLingvo, CCollLingvo)
OBJECT_ENTRY(CLSID_CompaundFiles, CCompaundFiles)
OBJECT_ENTRY(CLSID_Factor, CFactor)
OBJECT_ENTRY(CLSID_CollFac, CCollFac)
OBJECT_ENTRY(CLSID_MGertNet, CMGertNet)
OBJECT_ENTRY(CLSID_SafetyPrecaution, CSafetyPrecaution)
OBJECT_ENTRY(CLSID_FChange, CFChange)
OBJECT_ENTRY(CLSID_ICollFChange, CICollFChange)
OBJECT_ENTRY(CLSID_CollSF, CCollSF)
OBJECT_ENTRY(CLSID_CatRegistrar, CCatRegistrar)
END_OBJECT_MAP()

class CGertNetApp : public CWinApp
{
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGertNetApp)
	public:
    virtual BOOL InitInstance();
    virtual int ExitInstance();
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CGertNetApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CGertNetApp, CWinApp)
	//{{AFX_MSG_MAP(CGertNetApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

CGertNetApp theApp;

BOOL CGertNetApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
    hProxyDll = m_hInstance;
#endif
    _Module.Init(ObjectMap, m_hInstance, &LIBID_GERTNETLib);
    return CWinApp::InitInstance();
}

int CGertNetApp::ExitInstance()
{
    _Module.Term();
    return CWinApp::ExitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllCanUnloadNow() != S_OK)
        return S_FALSE;
#endif
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

    return (AfxDllCanUnloadNow()==S_OK && _Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#ifdef _MERGE_PROXYSTUB
    if (PrxDllGetClassObject(rclsid, riid, ppv) == S_OK)
        return S_OK;
#endif
    return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
#ifdef _MERGE_PROXYSTUB
    HRESULT hRes = PrxDllRegisterServer();
    if (FAILED(hRes))
        return hRes;
#endif
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _Module.RegisterServer( TRUE );
	if( SUCCEEDED(hr) )
	 {
	   CATEGORYINFO catInf[1];
	   catInf[ 0 ].catid = CATID_FacValMonitors;
	   catInf[ 0 ].lcid = MAKELCID( MAKELANGID(LANG_RUSSIAN,SUBLANG_DEFAULT), SORT_DEFAULT );
	   wcscpy( catInf[ 0 ].szDescription, L"Модули экспертной оценки значений факторов опасности" );   
      
	   CComPtr<ICatRegister> spCatReg;
	   hr = spCatReg.CoCreateInstance( CLSID_StdComponentCategoriesMgr );
	   if( SUCCEEDED(hr) )			
	     hr = spCatReg->RegisterCategories( sizeof(catInf)/sizeof(catInf[0]), catInf );
	 }

	return hr;
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
#ifdef _MERGE_PROXYSTUB
    PrxDllUnregisterServer();
#endif
    HRESULT hr = _Module.UnregisterServer( TRUE );

	if( SUCCEEDED(hr) )
	 {	   
       CComPtr<ICatRegister> spCatReg;
	   hr = spCatReg.CoCreateInstance( CLSID_StdComponentCategoriesMgr );
       
        CATID catID = (CATID)CATID_FacValMonitors;

	   if( SUCCEEDED(hr) )			
	     hr = spCatReg->UnRegisterCategories( 1, &catID );
	 }

	return hr;
}


