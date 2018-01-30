#if !defined(_INT_CREATOR_HELPER_)
#define _INT_CREATOR_HELPER_

#include "resource.h"       // main symbols

//#define MY_PASTER( ID ) L ###ID

template< 
  class SUBCLASS, 
  class INTF, 
  class COMCLS, 
  const IID* piidClient > class CPNCreator
 { 

public: 
	HRESULT __fastcall NewInst( SUBCLASS* pSub, INTF** pIntf, LPCWSTR lpwObjName, LPCWSTR lpwIntfName )  throw() 
	 { 
	   CComPtr<INTF> spOut;
	   CComObject<COMCLS>* pObj;	   
	   HRESULT hr;

	   try {
		  hr = CComObject<COMCLS>::CreateInstance( &pObj );
		  if( FAILED(hr) )		   
		    pSub->Error( (basic_string<WCHAR>(L"Cann't create ") + basic_string<WCHAR>( lpwObjName )).c_str(), *piidClient, hr );
		  else		  		  
		   {
			 hr = pObj->_InternalQueryInterface( __uuidof(INTF), (void**)&spOut );
			 if( FAILED(hr) )
			  {
				delete pObj;		  
				pSub->Error( (basic_string<WCHAR>(L"Cann't get ") + basic_string<WCHAR>( lpwIntfName ) + basic_string<WCHAR>(L" interface")).c_str(), *piidClient, hr );				
			  }
			 else
			  {
				CComPtr<IPersistStreamInit> spIPSI;
				hr = pObj->_InternalQueryInterface( IID_IPersistStreamInit, (void**)&spIPSI );
				if( FAILED(hr) )			
				  pSub->Error( L"Cann't get IPersistStreamInit", *piidClient, hr );
			    else				 
				  hr = spIPSI->InitNew();				   				 
			  }
		   }
		}
	   catch( bad_alloc e )
		{	  
		  hr = E_OUTOFMEMORY;
		}
	   catch( exception e )
		{	  
		  USES_CONVERSION;		  		  
		  hr = E_FAIL;
		  pSub->Error( A2COLE(e.what()), *piidClient, hr );		  
		}

	   if( FAILED(hr) ) *pIntf = NULL;
	   else *pIntf = spOut.Detach(); 

	   return hr;	
	 } 

	HRESULT __fastcall NewLoad( SUBCLASS* pSub, IStream* pStm, VARIANT_BOOL CreCached, INTF** pIntf, LPCWSTR lpwObjName, LPCWSTR lpwIntfName )  throw() 
	 { 
	   CComPtr<INTF> spOut;
	   CComObject<COMCLS> *pObj;
	   HRESULT hr;

	   try {
		  hr = CComObject<COMCLS>::CreateInstance( &pObj );
		  if( FAILED(hr) )		   
		    pSub->Error( (basic_string<WCHAR>(L"Cann't create ") + basic_string<WCHAR>( lpwObjName )).c_str(), *piidClient, hr );
		  else		  		  
		   {		     
			 hr = pObj->_InternalQueryInterface( __uuidof(INTF), (void**)&spOut );
			 if( FAILED(hr) )
			  {
				delete pObj;
				pSub->Error( (basic_string<WCHAR>(L"Cann't get ") + basic_string<WCHAR>( lpwIntfName ) + basic_string<WCHAR>(L" interface")).c_str(), *piidClient, hr );				
			  }
			 else
			  {
			    CComPtr<IStCollItem> spCItem;
				hr = pObj->_InternalQueryInterface( IID_IStCollItem, (void**)&spCItem );
				if( FAILED(hr) )
				   pSub->Error( L"Cann't get IStCollItem", *piidClient, hr );
				else
				 {
				   spCItem->put_DefaultCU( CreCached );
                 
				   CComPtr<IPersistStreamInit> spIPSI;
				   hr = pObj->_InternalQueryInterface( IID_IPersistStreamInit, (void**)&spIPSI );
				   if( FAILED(hr) )			
					 pSub->Error( L"Cann't get IPersistStreamInit", *piidClient, hr );
				   else				 
					 hr = spIPSI->Load( pStm );
				 }
			  }
		   }
		}
	   catch( bad_alloc e )
		{	  
		  hr = E_OUTOFMEMORY;
		}
	   catch( exception e )
		{	  
		  USES_CONVERSION;		  		  
		  hr = E_FAIL;
		  pSub->Error( A2COLE(e.what()), *piidClient, hr );		  
		}

	   if( FAILED(hr) ) *pIntf = NULL;
	   else *pIntf = spOut.Detach(); 

	   return hr;	
	 } 
 }; 

template< 
  class SUBCLASS, 
  class INTF, 
  class COMCLS, 
  const IID* piidClient > class CPNCreatorStg
 { 

public: 
	HRESULT __fastcall NewInst( SUBCLASS* pSub, IStorage* pStg, INTF** pIntf, LPCWSTR lpwObjName, LPCWSTR lpwIntfName )  throw() 
	 { 
	   *pIntf = NULL;
       if( pStg == NULL ) return E_POINTER;
	   CComPtr<INTF> spOut;

	   CComObject<COMCLS>* pObj;
	   HRESULT hr;

	   try {
		  hr = CComObject<COMCLS>::CreateInstance( &pObj );
		  if( FAILED(hr) )	
			pSub->Error( (basic_string<WCHAR>(L"Cann't create ") + basic_string<WCHAR>( lpwObjName )).c_str(), *piidClient, hr );		  
		  else	   
		   {	   
			 hr = pObj->_InternalQueryInterface( __uuidof(INTF), (void**)&spOut );
			 if( FAILED(hr) )
			  {
				delete pObj;		  
				pSub->Error( (basic_string<WCHAR>(L"Cann't get ") + basic_string<WCHAR>( lpwIntfName ) + basic_string<WCHAR>(L" interface")).c_str(), *piidClient, hr );			 
			  }
			 else
			  {
				CComPtr<IPersistStorage> spIPSI;
				hr = pObj->_InternalQueryInterface( IID_IPersistStorage, (void**)&spIPSI );
				if( FAILED(hr) )			   				
				  pSub->Error( L"Cann't get IPersistStorage", *piidClient, hr );
				else			  
				  hr = spIPSI->InitNew( pStg );			 
			  }
		   }
		}
	   catch( bad_alloc e )
		{	  
		  hr = E_OUTOFMEMORY;
		}
	   catch( exception e )
		{	  
		  USES_CONVERSION;		  		  
		  hr = E_FAIL;
		  pSub->Error( A2COLE(e.what()), *piidClient, hr );		  
		}

	   if( FAILED(hr) ) *pIntf = NULL;
	   else *pIntf = spOut.Detach(); 

	   return hr;	
	 } 

	HRESULT __fastcall NewLoad( SUBCLASS* pSub, IStorage* pStg, VARIANT_BOOL CreCached, INTF** pIntf, LPCWSTR lpwObjName, LPCWSTR lpwIntfName )  throw() 
	 { 
	   *pIntf = NULL;
       if( pStg == NULL ) return E_POINTER;
	   CComPtr<INTF> spOut;

	   CComObject<COMCLS>* pObj;
	   HRESULT hr;

	   try {
		  hr = CComObject<COMCLS>::CreateInstance( &pObj );
		  if( FAILED(hr) )	
			pSub->Error( (basic_string<WCHAR>(L"Cann't create ") + basic_string<WCHAR>( lpwObjName )).c_str(), *piidClient, hr );		  
		  else	   
		   {	   
			 hr = pObj->_InternalQueryInterface( __uuidof(INTF), (void**)&spOut );
			 if( FAILED(hr) )
			  {
				delete pObj;		  
				pSub->Error( (basic_string<WCHAR>(L"Cann't get ") + basic_string<WCHAR>( lpwIntfName ) + basic_string<WCHAR>(L" interface")).c_str(), *piidClient, hr );			 
			  }
			 else
			  {
			    CComPtr<IStCollItem> spCItem;
				hr = pObj->_InternalQueryInterface( IID_IStCollItem, (void**)&spCItem );
				if( FAILED(hr) )
				   pSub->Error( L"Cann't get IStCollItem", *piidClient, hr );
				else
				 {
				   spCItem->put_DefaultCU( CreCached );

				   CComPtr<IPersistStorage> spIPSI;
				   hr = pObj->_InternalQueryInterface( IID_IPersistStorage, (void**)&spIPSI );
				   if( FAILED(hr) )			   				
					 pSub->Error( L"Cann't get IPersistStorage", *piidClient, hr );
				   else			  
					 hr = spIPSI->Load( pStm ); 
				 }
			  }
		   }
		}
	   catch( bad_alloc e )
		{	  
		  hr = E_OUTOFMEMORY;
		}
	   catch( exception e )
		{	  
		  USES_CONVERSION;		  		  
		  hr = E_FAIL;
		  pSub->Error( A2COLE(e.what()), *piidClient, hr );		  
		}

	   if( FAILED(hr) ) *pIntf = NULL;
	   else *pIntf = spOut.Detach(); 

	   return hr;	
	 } 
 }; 

#endif

