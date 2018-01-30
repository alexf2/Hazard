// CatRegistrar.cpp : Implementation of CCatRegistrar
#include "stdafx.h"
#include "GertNet.h"


#include <list>
#include <algorithm>
using namespace std;

#include "CatRegistrar.h"


// {6E8BCCF5-6AE0-11d4-8FBA-00504E02C39D}
const CATID CATID_FacValMonitors = 
 { 0x6e8bccf5, 0x6ae0, 0x11d4, { 0x8f, 0xba, 0x0, 0x50, 0x4e, 0x2, 0xc3, 0x9d } };


/////////////////////////////////////////////////////////////////////////////
// CCatRegistrar

STDMETHODIMP CCatRegistrar::InterfaceSupportsErrorInfo(REFIID riid) throw()
{
	static const IID* arr[] = 
	{
		&IID_ICatRegistrar
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

static LPCWSTR warrPI[] =
  {
	L"Monitors.AmmoniaStoreHouse",
	L"Monitors.GasPipeLine",
	L"Monitors.RailWay"
  };

STDMETHODIMP CCatRegistrar::Register() throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif
   

   CATEGORYINFO catInf[1];
   catInf[ 0 ].catid = CATID_FacValMonitors;
   catInf[ 0 ].lcid = MAKELCID( MAKELANGID(LANG_RUSSIAN,SUBLANG_DEFAULT), SORT_DEFAULT );
   wcscpy( catInf[ 0 ].szDescription, L"Модули экспертной оценки значений факторов опасности" );   
      
   CComPtr<ICatRegister> spCatReg;
   HRESULT hr = spCatReg.CoCreateInstance( CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER );
   if( SUCCEEDED(hr) )	
	{
      hr = spCatReg->RegisterCategories( sizeof(catInf)/sizeof(catInf[0]), catInf );
	  if( SUCCEEDED(hr) )
	    for( int i = 0; i < sizeof(warrPI) / sizeof(warrPI[0]); ++i )
		 {
		   CLSID clsidTmp;
		   hr = CLSIDFromProgID( warrPI[i], &clsidTmp );
		   if( SUCCEEDED(hr) ) 
		     hr = spCatReg->RegisterClassImplCategories( clsidTmp, 1, &(CATID)CATID_FacValMonitors );

		   //if( FAILED() ) break;
		 }
	}
	
   return hr;
}

STDMETHODIMP CCatRegistrar::Unregister() throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


         
   CComPtr<ICatRegister> spCatReg;
   HRESULT hr = spCatReg.CoCreateInstance( CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER );
   if( SUCCEEDED(hr) )	
	{
      hr = spCatReg->UnRegisterCategories( 1, &(CATID)CATID_FacValMonitors );
	  if( SUCCEEDED(hr) )
	    for( int i = 0; i < sizeof(warrPI) / sizeof(warrPI[0]); ++i )
		 {
		   CLSID clsidTmp;
		   hr = CLSIDFromProgID( warrPI[i], &clsidTmp );
		   if( SUCCEEDED(hr) ) 
		     hr = spCatReg->UnRegisterClassImplCategories( clsidTmp, 1, &(CATID)CATID_FacValMonitors );

		   //if( FAILED() ) break;
		 }
	}
	
   return hr;
}

static void SFree( BSTR bs ) throw()
 {
   SysFreeString( bs );
 }

STDMETHODIMP CCatRegistrar::GetProgIDs( SAFEARRAY** arrProgIDs ) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif
 
    HRESULT hr = S_OK;

	try {
	   if( !arrProgIDs ) return E_POINTER;

	   list<BSTR> lstBS;

	   CComPtr<ICatInformation> spCatInf;
	   hr = spCatInf.CoCreateInstance( CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER );
	   if( SUCCEEDED(hr) )	
		{
		  CComPtr<IEnumCLSID> spEnum;
		  hr = spCatInf->EnumClassesOfCategories( 1, &(CATID)CATID_FacValMonitors, -1, NULL, &spEnum );
		  if( SUCCEEDED(hr) )
		   {
			 spEnum->Reset();
			 CLSID clsid;
			 while( (hr = spEnum->Next(1, &clsid, NULL)) == S_OK )
			  {
				LPWSTR lpwTmp;		     
				if( SUCCEEDED(ProgIDFromCLSID( clsid, &lpwTmp )) )
				 {
				   BSTR bsTmp_ = SysAllocString( lpwTmp );
				   CoTaskMemFree( lpwTmp );
				   if( bsTmp_ == NULL )
					{
					  hr = E_OUTOFMEMORY;
					  break;
					}
				   else
					 lstBS.push_back( bsTmp_ ); 			    
				 }			 
			  }		  
			 if( SUCCEEDED(hr) )		   
			   hr = CreArrrBSTR( lstBS, arrProgIDs );			 		   
			 else
			   Error( L"GetProgIDs: Cann't enumerate objects", IID_ICatRegistrar, hr );
			 
			 if( FAILED(hr) )
			   for_each( lstBS.begin(), lstBS.end(), SFree );

			 spEnum.Release();
		   }
		  else
			Error( L"GetProgIDs: Cann't EnumClassesOfCategories", IID_ICatRegistrar, hr );

		  spCatInf.Release();
		}
	   else	
		 Error( L"GetProgIDs: Cann't load cat manager", IID_ICatRegistrar, hr );
	 }
	catch( bad_alloc e )
	 {
	   hr = E_OUTOFMEMORY;
	 }
	catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( GetObjectCLSID(), A2COLE(e.what()), IID_ICatRegistrar, hr );
	}
	catch( CMemoryException *pE )
	 {
	   hr = E_OUTOFMEMORY;
	   pE->Delete();
	 }

	return hr;
}

HRESULT CCatRegistrar::CreArrrBSTR( list<BSTR>& rL, SAFEARRAY** ppArr ) throw()
 {

   SAFEARRAYBOUND bnd = { rL.size(), 0 };
   *ppArr = SafeArrayCreate( VT_BSTR, 1, &bnd );

   HRESULT hr = S_OK;

   if( !*ppArr )
	{
	  Error( L"Cann't array allocate", IID_ICatRegistrar, E_FAIL );
	  hr = E_FAIL;
	}
   else
	{
	  BSTR* pDta;
	  hr = SafeArrayAccessData( *ppArr, (void**)&pDta );
	  if( FAILED(hr) )
	   {
		 SafeArrayDestroy( *ppArr );
		 *ppArr = NULL;	   
		 Error( L"Cann't array access", IID_ICatRegistrar, hr );		 
	   }
	  else
	   {
		 copy( rL.begin(), rL.end(), pDta );
		 SafeArrayUnaccessData( *ppArr );
	   }
	}

   return hr;
 }
