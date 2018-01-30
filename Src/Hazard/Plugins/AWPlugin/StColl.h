#if !defined(_ST_COLL_)
#define _ST_COLL_


#include "resource.h"       // main symbols

//#pragma warning( push )
#pragma warning( disable: 4166 4503 )

struct lessBSTRNoCase: binary_function<CComBSTR, CComBSTR, bool> 
 {
   bool __fastcall operator()(const CComBSTR& _X, const CComBSTR& _Y) const throw()
	{
	  if( _Y.m_str == NULL && _X.m_str == NULL )
		  return false;
	  if( _Y.m_str != NULL && _X.m_str != NULL )
		  return _wcsicmp( _X.m_str, _Y.m_str ) < 0;
	  return _X.m_str == NULL;
	}
 };

__declspec(selectany) const wchar_t * lpwNAME_CONTENTS_STREAM = L"ObjContents";
__declspec(selectany) const wchar_t * lpwNO_NAME = L"<Нет>";


template<  
  class COMCLS, 
  const IID* pIID, 
  const CLSID* pCLSID > class ATL_NO_VTABLE CCreINI
 {
public:
   bool m_bDefaultCachedUpdates;

   CCreINI() throw()
	{
	  m_bDefaultCachedUpdates = false;
	}   

   HRESULT __fastcall CreateInstNotInit( IStCollItem** pIntf ) throw()
	{
	  CComObject<COMCLS>* pObj;
	  HRESULT hr;

	  try {
		 hr = CComObject<COMCLS>::CreateInstance( &pObj );
		 if( FAILED(hr) )
		   AtlReportError( *pCLSID, L"Cann't create object", *pIID, hr );
		 else
		  {
			hr = pObj->_InternalQueryInterface( __uuidof(IStCollItem), (void**)pIntf );
			if( FAILED(hr) )
			 {
			   delete pObj;		  
			   AtlReportError( *pCLSID, L"Cann't QI", *pIID, hr );		 
			 }
			else
			 {
			   VARIANT_BOOL vtB = (m_bDefaultCachedUpdates ? VARIANT_TRUE:VARIANT_FALSE);

			   (*pIntf)->put_DefaultCU( vtB );

			   CComDispatchDriver cdd( (IUnknown*)*pIntf );
			   if( cdd.p != NULL )
				{
				  CComVariant vStm, vMode;
				  V_VT(&vStm) = VT_DISPATCH, V_DISPATCH(&vStm) = NULL;
				  V_VT(&vMode) = VT_BOOL, V_BOOL(&vMode) = vtB;
				  cdd.Invoke2( L"SetUpdateMode", &vStm, &vMode );			
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
		 AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );		 
	   }

	  return hr;
	}
 };

template< class COLL, class COMCLS, const IID* pIID, const CLSID* pCLSID >
class ATL_NO_VTABLE CPersistToStreamImpl: 
  public CCreINI<COMCLS, pIID, pCLSID>
 {
public:
   bool __fastcall IsStreamPS() const throw() { return true; }
   STGTY __fastcall StorageFlag() const throw() { return STGTY_STREAM; }
   HRESULT __fastcall FetchItemName( IStorage* pStg, LPCWSTR lpName, CComBSTR& rBsRes ) throw()
	{
	  CComPtr<IStream> spStream;
	  HRESULT hr = pStg->OpenStream( lpName, 0,
	    STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spStream );

      if( SUCCEEDED(hr) )
	   {
	     long lTmp;
	     hr = spStream->Read( &lTmp, sizeof lTmp, NULL );
		 if( SUCCEEDED(hr) )		  
		   rBsRes.Empty(),
		   hr = rBsRes.ReadFromStream( spStream );					  
	   }

	  return hr;
	}

   HRESULT __fastcall LoadFrom( IStorage* pStm, LPCWSTR lpwName, CComPtr<IStCollItem>& rSp ) throw()
	{
	  HRESULT hr = CreateInstNotInit( &rSp );
      if( SUCCEEDED(hr) )
	   {
	     rSp->put_Sign( (long)static_cast<COLL*>(this)->GetUnknown() );

	     CComPtr<IPersistStreamInit> spIPSI;
		 hr = rSp.QueryInterface( &spIPSI );
		 if( FAILED(hr) )		  
		   AtlReportError( *pCLSID, L"Cann't QI IPersistStreamInit", *pIID, hr );					 
		 else
		  {
		    CComPtr<IStream> spStream;
			hr = pStm->OpenStream( lpwName, 0,
			  STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spStream );

			if( FAILED(hr) )			 	 		  	   
		      ReportStgError( hr, L"LoadFrom", *pCLSID, *pIID, lpwName );			 
			else			 			 
			  hr = spIPSI->Load( spStream );		 			   			 
		  }
	   }

	  if( FAILED(hr) ) rSp.Release();

	  return hr;
	}

   HRESULT __fastcall StoreExisted( IStorage* pStm, IStCollItem* pObj, DWORD dwKey ) throw()
	{
      CComPtr<IPersistStreamInit> spIPSI;
	  HRESULT hr = pObj->QueryInterface( IID_IPersistStreamInit, (void**)&spIPSI );
	  if( FAILED(hr) )		  
		AtlReportError( *pCLSID, L"Cann't QI IPersistStreamInit", *pIID, hr );					 
	  else
	   {
		 CComPtr<IStream> spStream;
		 wchar_t wcTmp[ 1024 ];
		 _ultow( dwKey, wcTmp, 10 );
		 hr = pStm->CreateStream( wcTmp,
		   STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spStream );

		 if( FAILED(hr) )			 	 		  	   
		   ReportStgError( hr, L"StoreExisted", *pCLSID, *pIID, wcTmp );
		 else			 			 
		   hr = spIPSI->Save( spStream, TRUE );
	   }

	  return hr;
	}

   HRESULT __fastcall StoreNew( IStorage* pStm, IStCollItem* pObj, DWORD dwKey, bool bReplace ) throw()
	{
	  return StoreExisted( pStm, pObj, dwKey );
	}

   HRESULT __fastcall CreateNew( IStorage*, DWORD dwKey, IStCollItem** pObj ) throw()
	{
	  CComPtr<IStCollItem> sp;
	  
	  HRESULT hr = CreateInstNotInit( &sp );
	  if( SUCCEEDED(hr) )
	   {
	     sp->put_Sign( (long)static_cast<COLL*>(this)->GetUnknown() );
		 sp->put_Key( dwKey );

	     CComPtr<IPersistStreamInit> spIPSI;
		 hr = sp.QueryInterface( &spIPSI );
		 if( FAILED(hr) )					   
		   AtlReportError( *pCLSID, L"Cann't get IPersistStreamInit", *pIID, hr );
		 else				 
		   hr = spIPSI->InitNew();				   				 
	   }

	  if( FAILED(hr) ) *pObj = NULL;
	  else *pObj = sp.Detach();

	  return hr;
	}
 };

//CPersistToStorageImpl 
template< class COLL, class COMCLS, const IID* pIID, const CLSID* pCLSID >
class ATL_NO_VTABLE CPersistToStorageImpl:
  public CCreINI<COMCLS, pIID, pCLSID>
 {
public:
   bool __fastcall IsStreamPS() const throw() { return false; }
   STGTY __fastcall StorageFlag() const throw() { return STGTY_STORAGE; }
   HRESULT __fastcall FetchItemName( IStorage* pStg, LPCWSTR lpName, CComBSTR& rBsRes ) throw()
	{
	  CComPtr<IStorage> spStorage;
	  HRESULT hr = pStg->OpenStorage( lpName, NULL, 
	    STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, NULL, 0, &spStorage );

      if( SUCCEEDED(hr) )
	   {
	     CComPtr<IStream> spStream;
		 hr = spStorage->OpenStream( lpwNAME_CONTENTS_STREAM, 0,
		   STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spStream );      

		 if( SUCCEEDED(hr) )
		  {
			long lTmp;
			hr = spStream->Read( &lTmp, sizeof lTmp, NULL );
			if( SUCCEEDED(hr) )		  
			  rBsRes.Empty(),
			  hr = rBsRes.ReadFromStream( spStream );					  
		  }
	   }

	  return hr;
	}

   HRESULT __fastcall LoadFrom( IStorage* pStm, LPCWSTR lpwName, CComPtr<IStCollItem>& rSp ) throw()
	{
	  HRESULT hr = CreateInstNotInit( &rSp );
      if( SUCCEEDED(hr) )
	   {
	     rSp->put_Sign( (long)static_cast<COLL*>(this)->GetUnknown() );

	     CComPtr<IPersistStorage> spIPSI;
		 hr = rSp->QueryInterface( &spIPSI );
		 if( FAILED(hr) )		  
		   AtlReportError( *pCLSID, L"Cann't QI IPersistStorage", *pIID, hr );					 
		 else
		  {
		    STATSTG stg;
		    hr = pStm->Stat( &stg, STATFLAG_NONAME );
			if( SUCCEEDED(hr) )
			 {
			   CComPtr<IStorage> spStorage;
			   hr = pStm->OpenStorage( lpwName, NULL,
				 stg.grfMode/*&(~STGM_CREATE)*//*STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE*/, 
				 0, 0, &spStorage );

			   if( FAILED(hr) )			 	 		  	   
				 ReportStgError( hr, L"LoadFrom", *pCLSID, *pIID, lpwName );			 
			   else			 			 
				 hr = spIPSI->Load( spStorage );		 			   			 
			 }
		  }
	   }

	  if( FAILED(hr) ) rSp.Release();

	  return hr;
	}

   HRESULT __fastcall StoreExisted( IStorage* pStm, IStCollItem* pObj, DWORD dwKey ) throw()
	{
      CComPtr<IPersistStorage> spIPSI;
	  HRESULT hr = pObj->QueryInterface( IID_IPersistStorage, (void**)&spIPSI );
	  if( FAILED(hr) )		  
		AtlReportError( *pCLSID, L"Cann't QI IPersistStorage", *pIID, hr );
	  else
	   {
		 CComPtr<IStorage> spStorage;
         
		 hr = pObj->get_Storage( &spStorage );
		 if( SUCCEEDED(hr) && spStorage.p == NULL )
		  {
			wchar_t wcTmp[ 1024 ];
			_ultow( dwKey, wcTmp, 10 );
			STATSTG stg;
			hr = pStm->Stat( &stg, STATFLAG_NONAME );
			if( SUCCEEDED(hr) )
			 {		    
			   hr = pStm->OpenStorage( wcTmp, NULL,
				 stg.grfMode/*&(~STGM_CREATE)*/, NULL, 0, &spStorage );

			   if( FAILED(hr) )			 	 		  	   
				 ReportStgError( hr, L"StoreExisted", *pCLSID, *pIID, wcTmp );		
			   else
				 hr = spIPSI->Save( spStorage, TRUE );			   
			 }
		  }		 
		 else if( SUCCEEDED(hr) )
		   hr = spIPSI->Save( NULL, TRUE );
	   }

	  return hr;
	}

   HRESULT __fastcall StoreNew( IStorage* pStm, IStCollItem* pObj, DWORD dwKey, bool bReplace ) throw()
	{
      CComPtr<IPersistStorage> spIPSI;
	  HRESULT hr = pObj->QueryInterface( IID_IPersistStorage, (void**)&spIPSI );
	  if( FAILED(hr) )		  
		AtlReportError( *pCLSID, L"Cann't QI IPersistStorage", *pIID, hr );
	  else
	   {
		 CComPtr<IStorage> spStorage;

		 hr = pObj->get_Storage( &spStorage );
		 if( SUCCEEDED(hr) && spStorage.p == NULL )
		  {
			wchar_t wcTmp[ 1024 ];
			_ultow( dwKey, wcTmp, 10 );
			STATSTG stg;
			hr = pStm->Stat( &stg, STATFLAG_NONAME );
			if( SUCCEEDED(hr) )
			 {
			   hr = pStm->CreateStorage( wcTmp,
				 !bReplace ? (stg.grfMode|STGM_CREATE|STGM_FAILIFTHERE):
				 (stg.grfMode|STGM_CREATE)&(~STGM_FAILIFTHERE),
				 0, 0, &spStorage );

			   if( FAILED(hr) )			 	 		  	   
				 ReportStgError( hr, L"StoreNew", *pCLSID, *pIID, wcTmp );
			   else					 
				 hr = spIPSI->Save( spStorage, FALSE );			   
			 }
		  }
		 else if( SUCCEEDED(hr) )
		   hr = spIPSI->Save( NULL, TRUE );
	   }

	  return hr;
	}

   HRESULT __fastcall CreateNew( IStorage* pStm, DWORD dwKey, IStCollItem** pObj ) throw()
	{
	  CComPtr<IStCollItem> sp;
	  
	  HRESULT hr = CreateInstNotInit( &sp );
	  if( SUCCEEDED(hr) )
	   {
	     sp->put_Sign( (long)static_cast<COLL*>(this)->GetUnknown() );
		 sp->put_Key( dwKey );

		 CComPtr<IPersistStorage> spIPS;
		 hr = sp.QueryInterface( &spIPS );

		 if( FAILED(hr) )		  			
		   AtlReportError( *pCLSID, L"Cann't QI IPersistStorage", *pIID, hr );
		 else
		  {
			if( pStm == NULL )
			  hr = spIPS->InitNew( NULL );
			else
			 {
			   wchar_t wcTmp[ 1024 ];
			   _ultow( dwKey, wcTmp, 10 );
			   STATSTG stg;
			   hr = pStm->Stat( &stg, STATFLAG_NONAME );
			   if( SUCCEEDED(hr) )
				{
				  CComPtr<IStorage> spStorage;

				  hr = pStm->CreateStorage( wcTmp,
					(stg.grfMode|STGM_CREATE)&(~STGM_FAILIFTHERE),
					0, 0, &spStorage );

				  if( FAILED(hr) )			 	 		  	   
					ReportStgError( hr, L"CreateNew", *pCLSID, *pIID, wcTmp );
				  else					 
				   {				  				  				  
					 hr = spIPS->InitNew( spStorage );
					 if( FAILED(hr) )
					  {
						spStorage.Release();
						pStm->DestroyElement( wcTmp );
					  }				  
				   }
				}
			 }
		  }
	   }

	  if( FAILED(hr) ) *pObj = NULL;
	  else *pObj = sp.Detach();

	  return hr;
	}
 };

#define PDSPV(a) V_DISPATCH(reinterpret_cast<VARIANT*>(const_cast<void*>(a)))
#define PSCI(a) (*reinterpret_cast<IStCollItem**>(const_cast<void*>(a)))

inline int __cdecl CmpKeyAsc( const void* a1, const void* a2 ) throw()
 {      
   if( PSCI(a1) == NULL && PSCI(a2) == NULL ) 
	 return 0;
   if( PSCI(a1) != NULL && PSCI(a2) != NULL )
	{
	  long k1, k2;
	  PSCI(a1)->get_Key( &k1 );
	  PSCI(a2)->get_Key( &k2 );
	  return k1 - k2;
	}
   return PSCI(a1) == NULL ? -1:1;
 }


inline int __cdecl CmpKeyDesc( const void* a1, const void* a2 ) throw()
 {      
   if( PSCI(a1) == NULL && PSCI(a2) == NULL ) 
	 return 0;
   if( PSCI(a1) != NULL && PSCI(a2) != NULL )
	{
	  long k1, k2;
	  PSCI(a1)->get_Key( &k1 );
	  PSCI(a2)->get_Key( &k2 );
	  return k2 - k1;
	}
   return PSCI(a1) == NULL ? 1:-1;
 }

inline int __cdecl CmpNameAsc( const void* a1, const void* a2 ) throw()
 {      
   if( PSCI(a1) == NULL && PSCI(a2) == NULL ) 
	 return 0;
   if( PSCI(a1) != NULL && PSCI(a2) != NULL )
	{
	  BSTR bs1, bs2;
	  bs1 = PSCI(a1)->NameByRef();
	  bs2 = PSCI(a2)->NameByRef();
	  return _wcsicmp( bs1, bs2 );
	}
   return PSCI(a1) == NULL ? -1:1;
 }

inline int __cdecl CmpNameDesc( const void* a1, const void* a2 ) throw()
 {      
   if( PSCI(a1) == NULL && PSCI(a2) == NULL ) 
	 return 0;
   if( PSCI(a1) != NULL && PSCI(a2) != NULL )
	{
	  BSTR bs1, bs2;
	  bs1 = PSCI(a1)->NameByRef();
	  bs2 = PSCI(a2)->NameByRef();
	  return _wcsicmp( bs2, bs1 );
	}
   return PSCI(a1) == NULL ? 1:-1;
 }


class _CopyIDispatch2Variant
 {
public:
	static HRESULT __fastcall copy( VARIANT* p1, IDispatch** p2 ) throw()
	 {
	   V_VT(p1) = VT_DISPATCH;
	   V_DISPATCH(p1) = *p2;
	   if( V_DISPATCH(p1) )
		 V_DISPATCH(p1)->AddRef();
	   return S_OK;
	 }
	static void __fastcall init2( VARIANT* p ) throw() { VariantInit(p); }
	static void __fastcall destroy2( VARIANT* p )  throw() { VariantClear(p); }
	static size_t __fastcall SizeNameBytes( IDispatch** ) throw()
	 {
	   return 11 * sizeof(wchar_t);
	 }
	static void __fastcall FormatName( IDispatch** pIt, LPWSTR lpw ) throw()
	 {
	   wcscpy( lpw, L"IDispatch*" );
	 }
 };


class _CopyLong2Variant
 {
public:
	static HRESULT __fastcall copy( VARIANT* p1, long* p2 ) throw()
	 {
	   V_VT(p1) = VT_I4;
	   V_I4(p1) = *p2;
	   
	   return S_OK;
	 }
	static void __fastcall init2( VARIANT* p ) throw() { VariantInit(p); }
	static void __fastcall destroy2( VARIANT* p )  throw() { VariantClear(p); }
	static size_t __fastcall SizeNameBytes( long *pl ) throw()
	 {
	   short i;

	   if( *pl == 0UL ) i = 1;
	   else
		{
		  __int64 i64L = (__int64)(ULONG)*pl;
		  for( i = 0; i64L >= 1; i64L /= 10, ++i );
		}

	   return size_t(i + 1) * sizeof(wchar_t);
	 }

	static void __fastcall FormatName( long* pIt, LPWSTR lpw ) throw()
	 {
	   _ultow( *pIt, lpw, 10 );
	 }
 };


template< 
  class T, 
  class PersistantImpl, 
  class Copy,
  class CopyOut,
  class ThreadModel = CComObjectThreadModel >
class ATL_NO_VTABLE CNCEnumImpl: 
  public CComObjectRootEx< ThreadModel >,
  public IEnumVARIANT,
  public PersistantImpl
 {
public:

   CNCEnumImpl() throw()
	{
	  m_bCachedUpdates = true;
	  m_begin = m_end = m_iter = NULL;
	}
   ~CNCEnumImpl() throw()
	{
	  for( T* p = m_begin; p != m_end; p++ )
		Copy::destroy( p );
	  delete [] m_begin;
	}

   STDMETHOD(Next)(ULONG celt, VARIANT* rgelt, ULONG* pceltFetched) throw();
   STDMETHOD(Skip)(ULONG celt) throw();
   STDMETHOD(Reset)(void) throw();
   STDMETHOD(Clone)(IEnumVARIANT** ppEnum) throw();

   HRESULT __fastcall Init(T* begin, T* end, IStorage* pStm = NULL, IUnknown* pUnk = NULL ) throw();

   bool GetCachedUpdates() const
	{
	  return m_bCachedUpdates;
	}
   void SetCachedUpdates( bool bFl ) 
	{
	  m_bCachedUpdates = bFl;
	}
   bool GetDefaultCachedUpdates() const
	{
	  return m_bDefaultCachedUpdates;
	}
   void SetDefaultCachedUpdates( bool bFl ) 
	{
	  m_bDefaultCachedUpdates = bFl;
	}

DECLARE_NOT_AGGREGATABLE(CNCEnumImpl)

BEGIN_COM_MAP(CNCEnumImpl)
  COM_INTERFACE_ENTRY(IEnumVARIANT)
END_COM_MAP()

  
	CComPtr<IUnknown> m_spUnk;
	CComPtr<IStorage> m_spStm;
	T* m_begin;
	T* m_end;
	T* m_iter;

	bool m_bCachedUpdates;
 };

template< 
  class T, 
  class PersistantImpl, 
  class Copy,
  class CopyOut,
  class ThreadModel >
HRESULT __fastcall CNCEnumImpl<T,PersistantImpl,Copy,CopyOut,ThreadModel>::Init(T* begin, T* end, IStorage* pStm, IUnknown* pUnk ) throw()
 {
   _ASSERTE( m_begin == NULL );
   m_iter = m_begin = begin;
   m_end = end;
   m_spUnk = pUnk;
   m_spStm = pStm;
   return S_OK;
 }

template< 
  class T, 
  class PersistantImpl, 
  class Copy,
  class CopyOut,
  class ThreadModel >
STDMETHODIMP CNCEnumImpl<T,PersistantImpl,Copy,CopyOut,ThreadModel>::Next( ULONG celt, VARIANT* rgelt, ULONG* pceltFetched ) throw()
 {
   if( rgelt == NULL || (celt != 1 && pceltFetched == NULL) )
	 return E_POINTER;

   HRESULT hr = S_OK;
   ULONG nActual = 0;
   VARIANT* pelt = rgelt;

   try {
	  size_t szBlock = 1024;
	  LPWSTR lpwName = NULL;
	  try { if( !m_bCachedUpdates ) lpwName = (LPWSTR)_alloca( szBlock ); }
	  catch(...)
	   {
	     hr = E_OUTOFMEMORY;				  
	   }
	  
	  while( SUCCEEDED(hr) && m_iter != m_end && nActual < celt )
	   {
         if( m_bCachedUpdates )
		  {
		    hr = CopyOut::copy( pelt, m_iter );
			if( FAILED(hr) )			 
			  while( rgelt < pelt ) CopyOut::destroy2( rgelt++ );			 
			else
			  nActual++, pelt++, m_iter++;
		  }
		 else
		  {
            _ASSERTE( m_spStm.p != NULL );

		    size_t sz = CopyOut::SizeNameBytes( m_iter );
			if( szBlock < sz )
			 {
			   szBlock = sz; lpwName = NULL;
			   try { lpwName = (LPWSTR)_alloca( szBlock ); }
			   catch(...)
				{
				  lpwName = NULL;
				}
			   if( lpwName == NULL )
				{
				  hr = E_OUTOFMEMORY;
				  break;
				}
			 }

			CopyOut::FormatName( m_iter, lpwName );

			CComPtr<IStCollItem> spTmpObj;			
			hr = LoadFrom( m_spStm, lpwName, spTmpObj );
            if( FAILED(hr) )
			  while( rgelt < pelt ) CopyOut::destroy2( rgelt++ );
			else
			 {			 
			   //VariantInit( pelt );
			   V_DISPATCH(pelt) = NULL;
			   hr = spTmpObj.QueryInterface( &V_DISPATCH(pelt) );
			   spTmpObj.Release();
			   if( FAILED(hr) )
			     while( rgelt < pelt ) CopyOut::destroy2( rgelt++ );
				else
				 {
				   V_VT(pelt) = VT_DISPATCH;
			       nActual++, pelt++, m_iter++;
				 }
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
	  AtlReportError( CLSID_NULL, A2COLE(e.what()), IID_IEnumVARIANT, hr );
	}
   

   if( pceltFetched )
     *pceltFetched = nActual;
   if( SUCCEEDED(hr) && (nActual < celt) )
	 hr = S_FALSE;

   return hr;
 }

template< 
  class T, 
  class PersistantImpl, 
  class Copy,
  class CopyOut,
  class ThreadModel >
STDMETHODIMP CNCEnumImpl<T,PersistantImpl,Copy,CopyOut,ThreadModel>::Skip( ULONG celt ) throw()
 {
   HRESULT hr;

   m_iter += celt;
   if( m_iter >= m_end )
    {
      m_iter = m_end;
      hr = S_FALSE;
    }
   else
   if( m_iter < m_begin )
    {	
      m_iter = m_begin;
      hr = S_FALSE;
    }
   else
	 hr = S_OK;

   return hr;
 }

template< 
  class T, 
  class PersistantImpl, 
  class Copy,
  class CopyOut,
  class ThreadModel >
STDMETHODIMP CNCEnumImpl<T,PersistantImpl,Copy,CopyOut,ThreadModel>::Reset( void ) throw()
 {
   m_iter = m_begin;
   return S_OK;
 }

template< 
  class T, 
  class PersistantImpl, 
  class Copy,
  class CopyOut,
  class ThreadModel >
STDMETHODIMP CNCEnumImpl<T,PersistantImpl,Copy,CopyOut,ThreadModel>::Clone( IEnumVARIANT** ppEnum ) throw()
 {
   typedef CComObject< CNCEnumImpl<T,PersistantImpl,Copy,CopyOut,ThreadModel> > _ThisClass;
   _ThisClass* pObj = NULL;
   HRESULT hr;
   CComPtr<IEnumVARIANT> spOut;
   bool bFlSta = false;
   
   try {
	  hr = _ThisClass::CreateInstance( &pObj );
      if( FAILED(hr) )
	     AtlReportError( CLSID_NULL, L"Cann't create object", IID_IEnumVARIANT, hr );
	  else
	   {
	     hr = pObj->_InternalQueryInterface( __uuidof(IEnumVARIANT), (void**)&spOut );
		 if( FAILED(hr) )
		  {
			delete pObj;		  			
			AtlReportError( CLSID_NULL, (basic_string<WCHAR>(L"Cann't get IEnumVARIANT") + basic_string<WCHAR>(L" interface")).c_str(), IID_IEnumVARIANT, hr );
		  }
		 else
		  {
		    pObj->m_spUnk = m_spUnk;
			pObj->m_spStm = m_spStm;
			pObj->m_bCachedUpdates = m_bCachedUpdates;			 

            pObj->m_begin = new T[ m_end - m_begin ];
			pObj->m_end = pObj->m_begin + (m_end - m_begin);
			pObj->m_iter = pObj->m_begin;			

			T* itSrc = m_begin;
			for( T* it = pObj->m_begin; pObj->m_iter != pObj->m_end; ++(pObj->m_iter), ++itSrc )
			 {
			   Copy::init( pObj->m_iter );
			   hr = Copy::copy( pObj->m_iter, itSrc );
			   if( FAILED(hr) ) break;
			   bFlSta = true;
			 }

            if( SUCCEEDED(hr) )			 
			  pObj->m_iter = pObj->m_begin + (m_iter - m_begin);			
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
	  AtlReportError( CLSID_NULL, A2COLE(e.what()), IID_IEnumVARIANT, hr );
	}

   if( FAILED(hr) ) 
	{
	  *ppEnum = NULL;
	  if( bFlSta )
	    for( --(pObj->m_iter); pObj->m_iter != pObj->m_begin; --(pObj->m_iter) )
		  Copy::destroy( pObj->m_iter );
 
	   if( pObj )
		{
		   if( pObj->m_begin ) delete[] pObj->m_begin;
		   pObj->m_begin = pObj->m_end = pObj->m_iter = NULL;
		}
	}
   else *ppEnum = spOut.Detach(); 

   return hr;
 }



template< 
  class COLL,
  class TYPE,   
  class BASE,  
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>

class ATL_NO_VTABLE CStColl: 
  public BASE, 
  public IPersistStorage,
  public IStCollItem,
  public PersistantImpl

 {
public:

  typedef map<DWORD, CComPtr<IStCollItem>, less<DWORD> > TMAP;
  typedef TMAP::iterator IT_TMAP;
  typedef TMAP::const_iterator CIT_TMAP;

  typedef map<CComBSTR, DWORD, lessBSTRNoCase > TMAP_DIR;
  typedef TMAP_DIR::iterator IT_TMAP_DIR;
  typedef TMAP_DIR::const_iterator CIT_TMAP_DIR;

  struct TFndByKey: public std::unary_function<TMAP_DIR::value_type&, bool>
   {
	 __fastcall TFndByKey( DWORD dw ) throw(): m_dw( dw )
	  {
	  }
	 bool __fastcall operator()( const TMAP_DIR::value_type& rObj ) throw()
	  {
	    return m_dw == rObj.second;
	  }
	DWORD m_dw;
   };


  CStColl() throw(): m_bsName(lpwNO_NAME) 
   {          
	 m_lKey = 0;
	 m_os = OS_None;  
	 m_lSign = 0;
	 m_lNewKey = 0;
	 m_sssSeq = SSS_AsIs;
	 
	 m_bCachedUpdates = false;	 	 
	 m_bInit = false;
   }
  virtual ~CStColl() throw()
   {
   }

  void __fastcall Modify() throw()
   {
     if( m_os == OS_None ) m_os = OS_Updated;
   }

//IPersistStorage
  STDMETHOD(GetClassID)( CLSID *pClassID ) throw()
	 {	   
	   return E_NOTIMPL;
	 }
  STDMETHOD(IsDirty)(void) throw()
   { 
     //if( !m_bCachedUpdates ) return S_FALSE;
     _ASSERTE( m_bInit );

	 if( m_os != OS_None ) return S_OK;
	 ObjStatus os;
	 for( CIT_TMAP it(m_mapObj.begin()); it != m_mapObj.end(); ++it )
	  { 	    
	    HRESULT hr = it->second->get_Status( &os );
		if( FAILED(hr) ) return hr;
		if( os != OS_None ) return S_OK;
	  }

	 return S_FALSE;
   }

  STDMETHOD(InitNew)( IStorage* pStorage ) throw();   
  STDMETHOD(Load)( IStorage* pStorage ) throw();   
  STDMETHOD(Save)( IStorage* pStorage, BOOL fSameAsLoad ) throw();
   
  STDMETHOD(SaveCompleted)(IStorage* /* pStorage */) throw()
   {
     return S_OK;
   }
  STDMETHOD(HandsOffStorage)(void) throw()
   {
     return S_OK;
   }

//Collection

  STDMETHOD(get__NewEnum)(/*[out, retval]*/ LPUNKNOWN *pVal) throw();
  STDMETHOD(Item)( /*[in]*/VARIANT* Key, /*[out, retval]*/TYPE **ppGT ) throw();
  STDMETHOD(get_Count)(/*[out, retval]*/ long *pVal) throw();
  STDMETHOD(Add)( /*[in]*/BSTR strName, /*[out,retval]*/TYPE** ppGT ) throw();
  STDMETHOD(Remove)(/*[in]*/VARIANT* KeyOrObj) throw();
  STDMETHOD(get_LongKey)(/*[in]*/BSTR Name, /*[out,retval]*/long* plKey) throw();
  STDMETHOD(get_NameKey)(/*[in]*/long Key, /*[in,optional,defaultvalue(0)]*/TYPE* Obj, /*[out,retval]*/BSTR* pbsName) throw();
  STDMETHOD(put_NameKey)(/*[in]*/long Key, /*[in,optional,defaultvalue(0)]*/TYPE* Obj,/*[in]*/BSTR bsNewName) throw();
  STDMETHOD(Clear)() throw();
  //STDMETHOD(ItemByName)(/*[in]*/BSTR Name, /*[out, retval]*/TYPE **ppGT) throw();
  STDMETHOD(Update)( /*[in]*/VARIANT* KeyOrObj ) throw();
  STDMETHOD(get_CachedUpdates)( /*[out,retval]*/VARIANT_BOOL* pbCached ) throw();
  //STDMETHOD(put_CachedUpdates)( /*[out,retval]*/VARIANT_BOOL pbCached ) throw(){return S_OK;}  
  STDMETHOD(SetUpdateMode)( IStorage* Stm, VARIANT_BOOL bMode ) throw();
  STDMETHOD(ItemNth)( /*[in]*/long Number, /*[out, retval]*/TYPE **ppGT ) throw();
  STDMETHOD(get_EnumSorting)( StdSortSeqv* pSortSeq ) throw();
  STDMETHOD(put_EnumSorting)( StdSortSeqv SortSeq ) throw();

  STDMETHOD(KeyNth)( /*[in]*/long Number, /*[out, retval]*/long* lKey ) throw();
  STDMETHOD(NameNth)( /*[in]*/long Number, /*[out, retval]*/BSTR* bsName ) throw();

  
  

//IStCollItem
  STDMETHOD(get_Key)( /*[out, retval]*/ long* plKey ) throw();
  STDMETHOD(put_Key)( /*[in]*/ long lKeyNew ) throw();
  STDMETHOD(get_Status)( /*[out, retval]*/ ObjStatus* pStatus ) throw();
  STDMETHOD(get_Name)( /*[out,retval]*/ BSTR* pbsName ) throw();
  STDMETHOD(put_Name)( /*[in]*/BSTR bsNameNew ) throw();
  STDMETHOD(put_Sign)( /*[in]*/long lSignNew ) throw();
  STDMETHOD(get_Sign)( /*[out,retval]*/long* plSign ) throw();
  STDMETHOD(get_Storage)( /*[out, retval]*/ IStorage** Stg ) throw();
  STDMETHOD(put_DefaultCU)( /*[in]*/VARIANT_BOOL bMode ) throw();
  STDMETHOD(get_DefaultCU)( /*[out,retval]*/VARIANT_BOOL* pbMode ) throw();
  

  /*[local]*/BSTR STDMETHODCALLTYPE NameByRef( void )
   {
	 return m_bsName;
   }

  
protected:
  long m_lKey;
  CComBSTR m_bsName;
  ObjStatus m_os;
  long m_lSign;
  DWORD m_lNewKey;
  StdSortSeqv m_sssSeq;
  

  CComPtr<IStorage> m_spStorage;
  bool m_bCachedUpdates;
  bool m_bInit;
  //bool m_bDefaultCachedUpdates;

  TMAP m_mapObj;
  TMAP_DIR m_mapDir;

  HRESULT __fastcall BuildDir( DWORD& rLMaxKey ) throw();
  HRESULT __fastcall InternalInitNew() throw();
  HRESULT __fastcall FullLoad() throw();
  HRESULT __fastcall FullSave() throw();  
  HRESULT __fastcall LoadContents() throw();
  HRESULT __fastcall SaveContents() throw();    
  HRESULT __fastcall KeyByName( BSTR bsName, DWORD& rdwKey, IT_TMAP_DIR* pIt = NULL ) throw();
  HRESULT __fastcall ExtractKeyFromIDispatch( IDispatch*, DWORD& rdwKey ) throw();
  void __fastcall ReportException( std::exception& e, HRESULT hr ) throw();
  DWORD __fastcall NextUniqueKey() throw();

  typedef int (__cdecl *TPF_CmpFunc)( const void*, const void* );
  TPF_CmpFunc GetCompareFunc_Objects() throw()
   {
     switch( m_sssSeq )
	  {
	    case SSS_AsIs:
		  return NULL;
	    case SSS_ByKeyAsc:
		  return CmpKeyAsc;
		case SSS_ByKeyDesc:
		  return CmpKeyDesc;
		case SSS_ByNameAsc:
		  return CmpNameAsc;
		case SSS_ByNameDesc:
		  return CmpNameDesc;
	  }
	 return NULL;
   }
  
  struct TItemManip
   {
	 __fastcall TItemManip( IStCollItem* p, DWORD dw ) throw()
	  {
		m_p = p; m_dw = dw;
	  }
	 IStCollItem* m_p;
	 DWORD m_dw;
   };
  typedef map<DWORD, TItemManip, less<DWORD> > MAP_TItemManip;
  typedef MAP_TItemManip::iterator IT_MAP_TItemManip;

  //typedef CComEnum<CComIEnum<VARIANT>, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT>, CComSingleThreadModel> _CComCollEnum;

  typedef CNCEnumImpl<
     IDispatch*,
	 PersistantImpl,
	 _CopyInterface<IDispatch>,
	 _CopyIDispatch2Variant,
	 CComObjectThreadModel > _CachedEnum;

  typedef CNCEnumImpl<
     long,
	 PersistantImpl,
	 _Copy<long>,
	 _CopyLong2Variant,
	 CComObjectThreadModel > _NotCachedEnum;

#ifdef _DEBUG
	void DumpLstRemov( list<DWORD>& ) throw();
	void DumpLstUpd( MAP_TItemManip& ) throw();
#endif
 };



template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::SaveContents()
 {
   HRESULT hr = S_OK;
   try {
	 if( m_spStorage.p == NULL )
	  {
	    hr = E_FAIL;
		AtlReportError( *pCLSID, L"SaveContents requires that the m_spStorage was not  NULL", *pIID, hr );
	  }
	 else
	  {
        CComPtr<IStream> spStream;
		hr = m_spStorage->CreateStream( lpwNAME_CONTENTS_STREAM, 
		  STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spStream );

		if( SUCCEEDED(hr) )
		 {		   
		   hr = spStream->Write( &m_lKey, sizeof m_lKey, NULL );
		   if( SUCCEEDED(hr) )		  
			 hr = m_bsName.WriteToStream( spStream );					  
		 }

		if( FAILED(hr) )		
		  ReportStgError( hr, L"SaveContents", *pCLSID, *pIID, lpwNAME_CONTENTS_STREAM );
	  }
	}
   catch( bad_alloc e )
	{	  
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
void __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::ReportException( std::exception& e, HRESULT hr ) throw()
 {
   USES_CONVERSION;
   AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::LoadContents() throw()
 {
   HRESULT hr = S_OK;
   try {
	 if( m_spStorage.p == NULL )
	  {
	    hr = E_FAIL;
		AtlReportError( *pCLSID, L"LoadContents requires that the m_spStorage was not  NULL", *pIID, hr );
	  }
	 else
	  {
        CComPtr<IStream> spStream;
		hr = m_spStorage->OpenStream( lpwNAME_CONTENTS_STREAM, 0,
		  STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spStream );

		if( SUCCEEDED(hr) )
		 {		   
		   hr = spStream->Read( &m_lKey, sizeof m_lKey, NULL );
		   if( SUCCEEDED(hr) )		  
			 m_bsName.Empty(), 
			 hr = m_bsName.ReadFromStream( spStream );					  
		 }

		if( FAILED(hr) )		
		  ReportStgError( hr, L"LoadContents", *pCLSID, *pIID, lpwNAME_CONTENTS_STREAM );
	  }
	}
   catch( bad_alloc e )
	{	  
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   return hr;
 }

#ifdef _DEBUG
template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
void CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::DumpLstRemov( list<DWORD>& rL ) throw()
 {   
   basic_stringstream<char> strm;

   list<DWORD>::iterator  it1( rL.begin() );
   
   strm << "Remove:" << endl;
   for( ; it1 != rL.end(); ++it1 )
     strm << *it1 << endl;

   OutputDebugString( strm.str().c_str() );
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
void CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::DumpLstUpd( MAP_TItemManip& rM ) throw()
 {
   USES_CONVERSION;
   basic_stringstream<char> strm;

   MAP_TItemManip::iterator  it1( rM.begin() );
      
   strm << "Update:" << endl;
   for( ; it1 != rM.end(); ++it1 )
	{
	  CComBSTR bs;
	  it1->second.m_p->get_Name( &bs );
      strm << it1->first << " - " << W2A(Chk(bs)) << endl;
	}

   OutputDebugString( strm.str().c_str() );
 }
#endif


template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::FullSave() throw()
 {
   HRESULT hr = S_OK;

   try {
	  if( m_spStorage.p == NULL )
	   {
		 hr = E_FAIL;
		 AtlReportError( *pCLSID, L"FullSave requires that the m_spStorage was not  NULL", *pIID, hr );
	   }
	  else
	   {
		 list<DWORD> lstRemov;
		 MAP_TItemManip lstUpdate;

		 CComPtr<IEnumSTATSTG> spEnum;
		 hr = m_spStorage->EnumElements( 0, NULL, 0, &spEnum );
		 if( SUCCEEDED(hr) )	  
		  {	     
			STATSTG stat;
			DWORD dwNum;
			while( spEnum->Next(1, &stat, NULL) == S_OK )
			 {
			   if( stat.type == StorageFlag() )
				{
				  if( stat.pwcsName == NULL || wcslen(stat.pwcsName) < 1 )
				   {
					 AtlReportError( *pCLSID, L"Bad item name in storage", *pIID, E_FAIL );
					 return E_FAIL;
				   }
				  wchar_t* lpwEnd;
				  dwNum = wcstoul( stat.pwcsName, &lpwEnd, 10 );
				  if( lpwEnd == stat.pwcsName )	
				   {
					 CoTaskMemFree( stat.pwcsName );
					 continue;
				   }						                  

				  IT_TMAP it = m_mapObj.find( dwNum );
				  if( it == m_mapObj.end() )
					lstRemov.push_back( dwNum );
				  else 
				   { 
					 ObjStatus os;
					 it->second->get_Status( &os );
					 if( os != OS_None ) lstUpdate.insert( MAP_TItemManip::value_type(dwNum, TItemManip(it->second, dwNum)) );
				   }
				}
			   if( stat.pwcsName ) CoTaskMemFree( stat.pwcsName ), stat.pwcsName = NULL;
			 }

			spEnum.Release();

#ifdef _DEBUG
	//__asm int 3h
	DumpLstRemov( lstRemov );
	DumpLstUpd( lstUpdate );
#endif

			list<DWORD>::const_iterator it( lstRemov.begin() );
			wchar_t wcTmp[ 1024 ];
			for( ; it != lstRemov.end(); ++it )
			 {
			   _ultow( *it, wcTmp, 10 );
			   if( FAILED(hr = m_spStorage->DestroyElement(wcTmp)) ) 
				{
				  ReportStgError( hr, L"FullSave", *pCLSID, *pIID, wcTmp );
				  return hr;
				}
			 }

			lstRemov.clear();

			IT_MAP_TItemManip it2( lstUpdate.begin() );
			for( ; it2 != lstUpdate.end(); ++it2 )
			 {
			   ObjStatus os = OS_None;
			   it2->second.m_p->get_Status( &os );
			   if( os != OS_New )
				 hr = StoreExisted(m_spStorage, it2->second.m_p, it2->second.m_dw );
			   else
				 hr = StoreNew( m_spStorage, it2->second.m_p, it2->second.m_dw, true );

			   if( FAILED(hr) )
				 {
				   _ultow( it2->second.m_dw, wcTmp, 10 );
				   ReportStgError( hr, L"FullSave", *pCLSID, *pIID, wcTmp );
				   return hr;
				 }
			 }

			CIT_TMAP cit( m_mapObj.begin() );
			for( ; cit != m_mapObj.end(); ++cit )
			 {
			   ObjStatus os = OS_None;
			   cit->second->get_Status( &os );
			   if( os != OS_None )
				{
				  it2 = lstUpdate.find( cit->first );
				  if( it2 == lstUpdate.end() )
					if( FAILED(hr = StoreNew(m_spStorage, cit->second.p, cit->first, false)) ) 
					 {
					   _ultow( cit->first, wcTmp, 10 );
					   ReportStgError( hr, L"FullSave", *pCLSID, *pIID, wcTmp );
					   return hr;
					 }
				}
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
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::FullLoad() throw()
 {
   HRESULT hr = S_OK;
   try {
	 if( m_spStorage.p == NULL )
	  {
	    hr = E_FAIL;
		AtlReportError( *pCLSID, L"FullLoad requires that the m_spStorage was not  NULL", *pIID, hr );
	  }
	 else
	  {
        m_mapDir.clear(); m_mapObj.clear();		
		m_lNewKey = 0;

		CComPtr<IEnumSTATSTG> spEnum;
		hr = m_spStorage->EnumElements( 0, NULL, 0, &spEnum );
		if( SUCCEEDED(hr) )	  
		 {	     
		   STATSTG stat;
		   DWORD dwNum;
		   while( spEnum->Next(1, &stat, NULL) == S_OK )
			{		    			  			  
			  if( stat.type == StorageFlag() )
			   {
			     if( stat.pwcsName == NULL || wcslen(stat.pwcsName) < 1 )
				  {
				    AtlReportError( *pCLSID, L"Bad item name in storage", *pIID, E_FAIL );
					return E_FAIL;
				  }
				 wchar_t* lpwEnd;
				 dwNum = wcstoul( stat.pwcsName, &lpwEnd, 10 );
				 if( lpwEnd == stat.pwcsName )
				  {
				    /*AtlReportError( *pCLSID, 
					   (basic_string<WCHAR>(L"Bad item name '") + 
					   basic_string<WCHAR>(stat.pwcsName) + 
					   basic_string<WCHAR>("' in storage - expected numeric")).c_str(), *pIID, E_FAIL );
					CoTaskMemFree( stat.pwcsName );
					return E_FAIL;*/
				    CoTaskMemFree( stat.pwcsName );
				    continue;
				  }
				 if( m_lNewKey < dwNum ) m_lNewKey = dwNum;

                 CComPtr<IStCollItem> spTmpObj;
				 hr = LoadFrom( m_spStorage, stat.pwcsName, spTmpObj );
				 if( FAILED(hr) )
				  {
				    ReportStgError( hr, L"FullLoad", *pCLSID, *pIID, stat.pwcsName );
				    CoTaskMemFree( stat.pwcsName );
					break;
				  }

                 
				 CComBSTR bsTmpName;
				 hr = spTmpObj->get_Name( &bsTmpName );
				 if( FAILED(hr) )
				  {
				    CoTaskMemFree( stat.pwcsName );
					break;
				  }
				 else
				  {
				    m_mapDir.insert( TMAP_DIR::value_type(bsTmpName, dwNum) );
					m_mapObj.insert( TMAP::value_type(dwNum, spTmpObj) );
					spTmpObj.Release();
				  }
                 
			   }
			  if( stat.pwcsName ) CoTaskMemFree( stat.pwcsName ), stat.pwcsName = NULL;
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
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::BuildDir( DWORD& rLMaxKey ) throw()
 {
   HRESULT hr = S_OK;
   try {
	 if( m_spStorage.p == NULL )
	  {
	    hr = E_FAIL;
		AtlReportError( *pCLSID, L"BuildDir requires that the m_spStorage was not  NULL", *pIID, hr );
	  }
	 else
	  {
        m_mapDir.clear();
		rLMaxKey = 0;

		CComPtr<IEnumSTATSTG> spEnum;
		hr = m_spStorage->EnumElements( 0, NULL, 0, &spEnum );
		if( SUCCEEDED(hr) )	  
		 {	     
		   STATSTG stg;
		   DWORD dwNum;
		   while( spEnum->Next(1, &stg, NULL) == S_OK )
			{		    			  			  
			  if( stg.type == StorageFlag() )
			   {
			     if( stg.pwcsName == NULL || wcslen(stg.pwcsName) < 1 )
				  {
				    AtlReportError( *pCLSID, L"Bad item name in storage", *pIID, E_FAIL );
					return E_FAIL;
				  }
				 wchar_t* lpwEnd;
				 dwNum = wcstoul( stg.pwcsName, &lpwEnd, 10 );
				 if( lpwEnd == stg.pwcsName )
				  {
				    /*AtlReportError( *pCLSID, 
					   (basic_string<WCHAR>(L"Bad item name '") + 
					   basic_string<WCHAR>(stat.pwcsName) + 
					   basic_string<WCHAR>("' in storage - expected numeric")).c_str(), *pIID, E_FAIL );
					CoTaskMemFree( stat.pwcsName );
					return E_FAIL;*/
				    CoTaskMemFree( stg.pwcsName );
				    continue;
				  }

				 CComBSTR bsTmp;
				 hr = FetchItemName( m_spStorage, stg.pwcsName, bsTmp );
				 if( FAILED(hr) )
				  {				    
					ReportStgError( hr, L"BuildDir", *pCLSID, *pIID, stg.pwcsName );	   
					CoTaskMemFree( stg.pwcsName );
					break;
				  }
				 else
				  {
				    m_mapDir.insert( TMAP_DIR::value_type(bsTmp, dwNum) );
					rLMaxKey = (rLMaxKey > dwNum ? rLMaxKey:dwNum);
				  }
			   }
			  if( stg.pwcsName ) CoTaskMemFree( stg.pwcsName ), stg.pwcsName = NULL;
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
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,  
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::InternalInitNew() throw()
 {
   HRESULT hr = S_OK;
   try {
	  m_mapObj.clear();
	  m_mapDir.clear();
	  m_os = OS_None;
	  
	  m_sssSeq = SSS_AsIs;	  	  	  
	  
	}
   catch( bad_alloc e )
	{	  
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::InitNew( IStorage* pStorage ) throw()
 {

   _ASSERTE( !m_bInit );

   HRESULT hr = S_OK;
   if( !m_bCachedUpdates && pStorage == NULL )
	{
	  hr = E_INVALIDARG;
	  AtlReportError( *pCLSID, L"In CachedUpdates mode need pStorage", *pIID, hr );
	}
   else
	{
      if( m_bCachedUpdates ) m_spStorage = NULL;

	  if( !::IsEmpty(pStorage) )
	   {
	     hr = E_FAIL;
	     AtlReportError( *pCLSID, L"Storage must been empty", *pIID, hr );
	   }
	  else
	   {	   
	     hr = InternalInitNew();
		 if( SUCCEEDED(hr) )
		  {
		    if( !m_bCachedUpdates ) m_spStorage = pStorage;         
		    m_os = OS_New;
			m_bInit = true;
		  }
	   }
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::Load( IStorage* pStorage ) throw()
 {
   _ASSERTE( !m_bInit );

   HRESULT hr = S_OK;
   
   if( pStorage == NULL ) 
	{
	  hr = E_POINTER;
	  AtlReportError( *pCLSID, L"pStorage must been not NULL", *pIID, hr );
	}
   else
	{
	  hr = InternalInitNew();	  
	  if( SUCCEEDED(hr) )
	   {
		 m_spStorage = pStorage;
		 m_os = OS_None;

		 hr = LoadContents();
		 if( SUCCEEDED(hr) )
		  {
			if( !m_bCachedUpdates )
			  hr = BuildDir( m_lNewKey );
			else
			  hr = FullLoad();

			if( SUCCEEDED(hr) ) ++m_lNewKey, m_bInit = true;
		  }
	   }
	}

   if( m_bCachedUpdates ) m_spStorage = NULL;
   
	
   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::Save( IStorage* pStorage, BOOL fSameAsLoad ) throw()
 {
   _ASSERTE( m_bInit );

   HRESULT hr = S_OK;

   if( pStorage == NULL && m_bCachedUpdates ) hr = E_POINTER;	
   else if( pStorage != NULL && !m_bCachedUpdates ) hr = E_POINTER;	
   else if( m_bCachedUpdates )
	{	  	  
	  m_spStorage = pStorage;
	  hr = SaveContents();
      if( SUCCEEDED(hr) ) hr = FullSave();		 		 	   
	}
   else
	{
	  _ASSERTE( m_spStorage.p != NULL );
	  //m_spStorage = pStorage; 
      hr = SaveContents();	     	   
	}

   HRESULT hr2 = EndTransaction( m_spStorage, hr );   
   if( FAILED(hr2) && SUCCEEDED(hr) ) 
	 ReportStgError( hr2, L"Save", *pCLSID, *pIID ),
	 hr = hr2;
   
	
   if( m_bCachedUpdates ) m_spStorage = NULL;
   //m_spStorage = NULL;
   if( SUCCEEDED(hr) ) m_os = OS_None;	  
   	
   return hr;
 }



template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_Storage( IStorage** Stg ) throw()
 {
   if( !Stg ) return E_POINTER;
   *Stg = m_spStorage.p;
   if( *Stg ) (*Stg)->AddRef();
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_Key( /*[out, retval]*/ long* plKey ) throw()
 {
   MGET_PROPERTY( plKey, m_lKey );
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::put_Key( /*[in]*/ long lKeyNew ) throw()
 {
   MPUT_PROPERTY( m_lKey, lKeyNew );
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_Status( /*[out, retval]*/ ObjStatus* pStatus ) throw()
 {
   //MGET_PROPERTY( pStatus, m_os );
   if( m_os != OS_None ) *pStatus = m_os;
   else if( IsDirty() == S_OK ) *pStatus = OS_Updated;
   else *pStatus = OS_None;

   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_Name( /*[out,retval]*/ BSTR* pbsName ) throw()
 {
   return m_bsName.CopyTo( pbsName );
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::put_Name( /*[in]*/BSTR bsNameNew ) throw()
 {
   if( m_bsName == bsNameNew ) return S_OK;

   m_bsName = bsNameNew;
   Modify();
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::put_Sign( /*[in]*/long lSignNew ) throw()
 {   
   m_lSign = lSignNew;
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_Sign( /*[out,retval]*/long* plSign ) throw()
 {
   MGET_PROPERTY( plSign, m_lSign );
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::Item( VARIANT* Key, TYPE **ppGT ) throw()
 {   
   if( Key == NULL ) return E_POINTER;   
   VARTYPE vt = (V_VT(Key) & VT_TYPEMASK);

   DWORD dwKey;
   HRESULT hr = S_OK;
   if( V_ISBYREF(Key) == VT_BYREF )
	{
	  switch( vt )
	   {
	     case VT_I2:
		   dwKey = *V_I2REF( Key );
		   break;
	     case VT_I4:
		   dwKey = *V_I4REF( Key );
		   break;
		 case VT_UI4:
		   dwKey = *V_UI4REF( Key );
		   break;
		 case VT_BSTR:
		  {
		    hr = KeyByName( *V_BSTRREF(Key), dwKey );
		    break;
		  }		 
		 default:
		   hr = E_INVALIDARG;
		   AtlReportError( *pCLSID, L"Bad type of KeyOrObj", *pIID, hr );
	   };
	}
   else
	{
	  switch( vt )
	   {
	     case VT_I2:
		   dwKey = V_I2( Key );
		   break;
	     case VT_I4:
		   dwKey = V_I4( Key );
		   break;
		 case VT_UI4:
		   dwKey = V_UI4( Key );
		   break;
		 case VT_BSTR:
		  {
		    hr = KeyByName( V_BSTR(Key), dwKey );
		    break;
		  }		 
		 default:
		   hr = E_INVALIDARG;
		   AtlReportError( *pCLSID, L"Bad type of KeyOrObj", *pIID, hr );
	   };
	}

   if( FAILED(hr) ) return hr;

   if( m_bCachedUpdates )
	{
	  IT_TMAP it = m_mapObj.find( dwKey );
	  if( it == m_mapObj.end() ) 
	   {
	     hr = E_FAIL;
		 basic_stringstream<WCHAR> strm;
		 strm << L"Item " << dwKey << L" is not found";
	     AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	   }
	  else
	   {
	     hr = it->second->QueryInterface( __uuidof(TYPE), (void**)ppGT );
		 if( FAILED(hr) )
		  {
		    basic_stringstream<WCHAR> strm;
			strm << L"Item " << dwKey << L" is found, but not supported requested interface";
			AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
		  }
	   }
	}
   else
	{
	  TFndByKey dta( dwKey );
	  IT_TMAP_DIR it = find_if( m_mapDir.begin(), m_mapDir.end(), dta );
      if( it == m_mapDir.end() ) 
	   {
	     hr = E_FAIL;
		 basic_stringstream<WCHAR> strm;
		 strm << L"Item " << dwKey << L" is not found";
	     AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	   }
	  else
	   {
         CComPtr<IStCollItem> spTmpObj;
		 wchar_t wcTmp[ 1024 ];
		 _ultow( dwKey, wcTmp, 10 );
		 hr = LoadFrom( m_spStorage, wcTmp, spTmpObj );
		 if( SUCCEEDED(hr) )
		  {
			hr = spTmpObj->QueryInterface( __uuidof(TYPE), (void**)ppGT );
			if( FAILED(hr) )
			 {
			   basic_stringstream<WCHAR> strm;
			   strm << L"Item " << dwKey << L" is found, but not supported requested interface";
			   AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
			 }
		  }
	   }
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_Count(/*[out, retval]*/ long *pVal) throw()
 {
   if( !pVal ) return E_POINTER;
   *pVal = m_mapDir.size();
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_LongKey(/*[in]*/BSTR Name, /*[out,retval]*/long* plKey) throw()
 {
   HRESULT hr = S_OK;

   CComBSTR bsTmp; bsTmp.Attach( Name );

   if( Name == NULL || plKey == NULL ) hr = E_POINTER;
   else
	{   	  	  
	  DWORD dwKey;
	  hr = KeyByName( bsTmp, dwKey );
	  if( SUCCEEDED(hr) )	     
		 *plKey = dwKey;	   
	}

   bsTmp.Detach();

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_NameKey( long Key, TYPE* Obj, BSTR* pbsName ) throw()
 {
   HRESULT hr = S_OK;

   if( Obj ) return E_POINTER;
   

   if( m_bCachedUpdates )
	{
	  IT_TMAP it = m_mapObj.find( Key );
	  if( it == m_mapObj.end() ) 
	   {
	     hr = E_FAIL;
		 basic_stringstream<WCHAR> strm;
		 strm << L"Item " << Key << L" is not found";
	     AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	   }
	  else	   
	    hr = it->second->get_Name( pbsName );	        
	}
   else
	{
	  TFndByKey dta( Key );
	  IT_TMAP_DIR it = find_if( m_mapDir.begin(), m_mapDir.end(), dta );
      if( it == m_mapDir.end() ) 
	   {
	     hr = E_FAIL;
		 basic_stringstream<WCHAR> strm;
		 strm << L"Item " << Key << L" is not found";
	     AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	   }
	  else
	   {
	     CComBSTR* pbsrTmp = const_cast<CComBSTR*>( &it->first );
	     hr = pbsrTmp->CopyTo( pbsName );
	   }
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::put_NameKey( long Key, TYPE* Obj, BSTR bsNewName ) throw()
 {
   HRESULT hr = S_OK;
      
   
   if( Obj )
	{
	  CComQIPtr<IStCollItem> spItTmp( Obj );
	  if( spItTmp.p == NULL )
	   {
	     AtlReportError( *pCLSID, L"Obj не поддерживает IStCollItem", *pIID, E_NOINTERFACE );
		 return E_NOINTERFACE;
	   }
	  hr = spItTmp->put_Name( bsNewName );
	  if( FAILED(hr) ) return hr;	  
	}

   CComBSTR bsNewN; bsNewN.Attach( bsNewName );

   TFndByKey dta( Key );
   IT_TMAP_DIR it = find_if( m_mapDir.begin(), m_mapDir.end(), dta );
   IT_TMAP_DIR it2 = m_mapDir.find( bsNewN );

   if( it == m_mapDir.end() )
	{
	  hr = E_FAIL;
	  basic_stringstream<WCHAR> strm;
	  strm << L"Item " << Key << L" is not found";
	  AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	}
   else
   if( it2 != m_mapDir.end() && it2 != it )
	{
	  hr = E_FAIL;
	  basic_stringstream<WCHAR> strm;
	  strm << L"Can't rename item '" << Chk(it->first) << L"'. Item '" << Chk(bsNewN) << L"' already exists";
	  AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	}      
   else
	{
      if( m_bCachedUpdates )
	   {
	     IT_TMAP it = m_mapObj.find( Key );
		 if( it == m_mapObj.end() ) 
		  {
			hr = E_FAIL;
			basic_stringstream<WCHAR> strm;
			strm << L"Item " << Key << L" is not found";
			AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
		  }
		 else	   		 		  
		   hr = it->second->put_Name( bsNewN );
	   }
	  else
	   {
	     CComPtr<IStCollItem> spTmpObj;
		 wchar_t wcTmp[ 1024 ];
		 _ultow( Key, wcTmp, 10 );

		 bool bKey = m_bDefaultCachedUpdates;
		 m_bDefaultCachedUpdates = false;

		 bool bUseObj;
		 if( (bUseObj = (Obj != NULL)) )
		   hr = Obj->QueryInterface( IID_IStCollItem, (void**)&spTmpObj );
		 else
		   hr = LoadFrom( m_spStorage, wcTmp, spTmpObj );

		 if( SUCCEEDED(hr) )
		  {
			if( !bUseObj ) hr = spTmpObj->put_Name( bsNewN );
			if( SUCCEEDED(hr) )
			 {
			   hr = StoreExisted( m_spStorage, spTmpObj, Key );			  

			   HRESULT hr2 = EndTransaction( m_spStorage, hr );   
			   if( FAILED(hr2) && SUCCEEDED(hr) ) 
				 ReportStgError( hr2, L"put_NameKey", *pCLSID, *pIID ),
				 hr = hr2;
			 }
		  }

		 m_bDefaultCachedUpdates = bKey;
	   }
	}


   if( SUCCEEDED(hr) )
	{
	  DWORD dwTmp = it->second;
	  m_mapDir.erase( it );
	  m_mapDir.insert( TMAP_DIR::value_type(bsNewN, dwTmp) );
	  if( m_bCachedUpdates ) Modify();
	}
   
   bsNewN.Detach();

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::put_DefaultCU( /*[in]*/VARIANT_BOOL bMode ) throw()
 {
   m_bDefaultCachedUpdates = (bMode == VARIANT_TRUE ? true:false);
   //m_bCachedUpdates = m_bDefaultCachedUpdates;
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_DefaultCU( /*[out,retval]*/VARIANT_BOOL* pbMode ) throw()
 {
   *pbMode = (m_bDefaultCachedUpdates ? VARIANT_TRUE:VARIANT_FALSE);
   return S_OK;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::Clear() throw()
 {   
   HRESULT hr = S_OK;

   try {
	  if( m_bCachedUpdates )
		m_mapObj.clear(),
		m_mapDir.clear();
	  else
	   {		 
		 IT_TMAP_DIR  it( m_mapDir.begin() );
		 for( ; it != m_mapDir.end(); ++it )
		  {		 
			wchar_t wcTmp[ 1024 ];
			
			_ultow( it->second, wcTmp, 10 );
			if( FAILED(hr = m_spStorage->DestroyElement(wcTmp)) ) 
			 {			   
			   ReportStgError( hr, L"Clear", *pCLSID, *pIID, wcTmp );
			   //if( it != m_mapDir.begin() ) m_mapDir.erase( m_mapDir.begin(), it );
			   break;
			 }
		  }		 
		 if( SUCCEEDED(hr) ) m_mapDir.clear();
	   }  	  
	}
   catch( bad_alloc e )
	{	  
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   if( !m_bCachedUpdates )
	{
	  HRESULT hr2 = EndTransaction( m_spStorage, hr );   
	  if( FAILED(hr2) && SUCCEEDED(hr) ) 
		ReportStgError( hr2, L"Clear", *pCLSID, *pIID ),
		hr = hr2;
	}

   if( m_bCachedUpdates ) Modify();

   return hr;
 }

/*template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::ItemByName(BSTR Name, TYPE **ppGT) throw()
 {
   HRESULT hr = S_OK;

   CComBSTR bsTmp; bsTmp.Attach( Name );

   if( ppGT == NULL ) hr = E_POINTER;
   else
	{   	  
	  DWORD dwK;
	  hr = KeyByName( bsTmp, dwK );
	  if( SUCCEEDED(hr) )
	   {	     		 
		 if( m_bCachedUpdates )
		  {
		    CIT_TMAP it2 = m_mapObj.find( dwK );
			if( it2 == m_mapObj.end() )
			 {
			   hr = E_FAIL;
			   AtlReportError( *pCLSID, L"Internal error", *pIID, hr );
			 }
			else			  
			  hr = it2->second->QueryInterface( __uuidof(TYPE), (void**)ppGT );
		  }
		 else
		  {
		    CComPtr<IStCollItem> spTmpObj;
			wchar_t wcTmp[ 1024 ];
			_ultow( dwK, wcTmp, 10 );
			hr = LoadFrom( m_spStorage, wcTmp, spTmpObj );
			if( SUCCEEDED(hr) )
			 {
			   hr = spTmpObj->QueryInterface( __uuidof(TYPE), (void**)ppGT );
			   if( FAILED(hr) )
				{
				  basic_stringstream<WCHAR> strm;
				  strm << L"Item '" << Chk(bsTmp) << L"' is found, but not supported requested interface";
				  AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
				}
			 }
		  }
	   }
	}

   bsTmp.Detach();
   return hr;
 }*/

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_CachedUpdates( /*[out,retval]*/VARIANT_BOOL* pbCached ) throw()
 {
   *pbCached = (m_bCachedUpdates ? VARIANT_TRUE:VARIANT_FALSE);
   return S_OK;
 }



template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::SetUpdateMode( IStorage* Stm, VARIANT_BOOL bMode ) throw()
 {
//__asm int 3h
   if( !m_bInit )
	{
	  m_bCachedUpdates = (bMode == VARIANT_TRUE ? true:false);
	  return S_OK;
	}

   HRESULT hr = S_OK;
   
   if( m_bCachedUpdates )
	{
	  if( bMode == VARIANT_TRUE )
	   {
	     if( Stm != NULL )
		  {
		    hr = E_FAIL;
			AtlReportError( *pCLSID, L"Can't bind to storage in CahchedUpdate mode", *pIID, hr );
		  }
	   }
	  else
	   {
	     if( IsDirty() == S_OK )
		  {
		    hr = E_FAIL;
			AtlReportError( *pCLSID, L"Can't set NotCahchedUpdate mode - changes not saved", *pIID, hr );
		  }
		 else
		  {
		    if( Stm == NULL )
			 {
			   hr = E_INVALIDARG;
			   AtlReportError( *pCLSID, L"Can't set NotCahchedUpdate mode - need storage", *pIID, hr );
			 }
			else
			 {
			   m_mapObj.clear();
			   m_bCachedUpdates = false;
			   m_spStorage = Stm;
			   hr = Load( m_spStorage );			
			 }
		  }
	   }
	}
   else
	{
	  if( bMode == VARIANT_TRUE )
	   {
	     if( Stm != NULL )
		  {
		    hr = E_FAIL;
			AtlReportError( *pCLSID, L"Can't bind to storage in CahchedUpdate mode", *pIID, hr );
		  }
		 else
		  {		    
			m_bCachedUpdates = true;
			hr = Load( m_spStorage );			
			m_spStorage = NULL;
		  }
	   }
	  else
	   {
	     if( Stm == NULL )
		  {
		    hr = E_FAIL;
			AtlReportError( *pCLSID, L"Can't unbind to storage in CahchedUpdate mode", *pIID, hr );
		  }
		 else		  
		  {
		    m_spStorage = NULL;
			hr = Load( Stm );			
		  }
	   }
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::Add( BSTR strName, TYPE** ppGT ) throw()
 {
   HRESULT hr = S_OK;
   CComBSTR bsName; bsName.Attach( strName );
//_asm int 3h
   DWORD dwKey;
   bool bObjInserted = false;
   bool bDirInserted = false;

   try {

	  IT_TMAP_DIR it = m_mapDir.find( bsName );
	  if( it != m_mapDir.end() )
	   {
		 hr = E_FAIL;
		 basic_stringstream<WCHAR> strm;
		 strm << L"Can't add new: item '" << Chk(bsName) << L"' already exists";
		 AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );	  
	   }
	  else
	   {
		 CComPtr<IStCollItem> spSCI;
		 //_asm  int 3h
		 hr = CreateNew( m_spStorage, dwKey=NextUniqueKey(), &spSCI );
		 
		 if( SUCCEEDED(hr) && SUCCEEDED(hr = spSCI.QueryInterface(ppGT)) )
		  {	     
		    //spSCI->put_Sign( (long)static_cast<COLL*>(this)->GetUnknown() );

			hr = spSCI->put_Name( strName );
			if( SUCCEEDED(hr) )
			 {			   
			   //hr = spSCI->put_Key( (dwKey=NextUniqueKey()) );
			   
			   if( m_bCachedUpdates )				   
				 bObjInserted = m_mapObj.insert( TMAP::value_type(dwKey, spSCI) ).second;
			   else				   
				{
				  hr = StoreNew( m_spStorage, spSCI, dwKey, false );

				  HRESULT hr2 = EndTransaction( m_spStorage, hr );   
				  if( FAILED(hr2) && SUCCEEDED(hr) ) 
					ReportStgError( hr2, L"Add", *pCLSID, *pIID ),
					hr = hr2;
				}
				

			   if( SUCCEEDED(hr) )								   
				 bDirInserted = m_mapDir.insert( TMAP_DIR::value_type(bsName, dwKey) ).second;
			
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
	  //USES_CONVERSION;
	  hr = E_FAIL;
	  //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	  ReportException( e, hr );
	}

   
   if( SUCCEEDED(hr) && m_bCachedUpdates ) Modify();
   if( FAILED(hr) )
	{
	  if( ppGT && *ppGT ) 
	    (*ppGT)->Release(), *ppGT = NULL;

	  if( m_bCachedUpdates )
	    if( bObjInserted ) m_mapObj.erase( dwKey );

	  if( bDirInserted ) m_mapDir.erase( bsName );
	}

   bsName.Detach();
   return hr;
 }



template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
DWORD __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::NextUniqueKey() throw()
 {
   if( m_bCachedUpdates )
	{
	  CIT_TMAP it;
	  while( (it = m_mapObj.find(m_lNewKey)) != m_mapObj.end() ) 
		m_lNewKey++;
	}
   else
	{
	  HRESULT hr;
	  do {
		 TFndByKey dta( m_lNewKey );
		 IT_TMAP_DIR it;
		 while( (it = find_if(m_mapDir.begin(), m_mapDir.end(), dta)) != m_mapDir.end() ) 
		   dta.m_dw = ++m_lNewKey;
	   } while( (hr = ObjExists(m_spStorage, m_lNewKey, StorageFlag())) == S_OK );
	}

   return m_lNewKey++;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::Remove( VARIANT* KeyOrObj ) throw()
 {
   if( KeyOrObj == NULL ) return E_POINTER;   
   VARTYPE vt = (V_VT(KeyOrObj) & VT_TYPEMASK);

   DWORD dwKey;
   HRESULT hr = S_OK;
   if( V_ISBYREF(KeyOrObj) == VT_BYREF )
	{
	  switch( vt )
	   {
	     case VT_I4:
		   dwKey = *V_I4REF( KeyOrObj );
		   break;
		 case VT_UI4:
		   dwKey = *V_UI4REF( KeyOrObj );
		   break;
		 case VT_BSTR:
		  {
		    hr = KeyByName( *V_BSTRREF(KeyOrObj), dwKey );
		    break;
		  }
		 case VT_DISPATCH:
		    hr = ExtractKeyFromIDispatch( *V_DISPATCHREF(KeyOrObj), dwKey );
			break;
		 default:
		   hr = E_INVALIDARG;
		   AtlReportError( *pCLSID, L"Bad type of KeyOrObj", *pIID, hr );
	   };
	}
   else
	{
	  switch( vt )
	   {
	     case VT_I4:
		   dwKey = V_I4( KeyOrObj );
		   break;
		 case VT_UI4:
		   dwKey = V_UI4( KeyOrObj );
		   break;
		 case VT_BSTR:
		  {
		    hr = KeyByName( V_BSTR(KeyOrObj), dwKey );
		    break;
		  }
		 case VT_DISPATCH:
		    hr = ExtractKeyFromIDispatch( V_DISPATCH(KeyOrObj), dwKey );
			break;
		 default:
		   hr = E_INVALIDARG;
		   AtlReportError( *pCLSID, L"Bad type of KeyOrObj", *pIID, hr );
	   };
	}

   if( SUCCEEDED(hr) )
	{
	  TFndByKey dta( dwKey );
	  IT_TMAP_DIR it = find_if( m_mapDir.begin(), m_mapDir.end(), dta );
	  if( it == m_mapDir.end() )
	   {
	     hr = E_FAIL;
		 AtlReportError( *pCLSID, L"Remove: Internal error", *pIID, hr );
	   }
	  else
	   {
		if( m_bCachedUpdates )
		 {
		   m_mapDir.erase( it );
		   m_mapObj.erase( dwKey );
		   Modify();
		 }
		else
		 {
		   wchar_t wcTmp[ 1024 ];
		   _ultow( dwKey, wcTmp, 10 );
		   if( FAILED(hr = m_spStorage->DestroyElement(wcTmp)) ) 
			 ReportStgError( hr, L"Remove", *pCLSID, *pIID, wcTmp );
		   else
			{
			  HRESULT hr2 = EndTransaction( m_spStorage, hr );   
			  if( FAILED(hr2) && SUCCEEDED(hr) ) 
				ReportStgError( hr2, L"Remove", *pCLSID, *pIID ),
				hr = hr2;

			  if( SUCCEEDED(hr) ) m_mapDir.erase( it );
			}
		 }
	   }
	}

   return hr;
 }



template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::KeyByName( BSTR bsName, DWORD& rdwKey, IT_TMAP_DIR* pIt ) throw()
 {
   HRESULT hr = S_OK;

   CComBSTR bsTmp; bsTmp.Attach( bsName );
   IT_TMAP_DIR it = m_mapDir.find( bsTmp );
   
   if( it == m_mapDir.end() )
	{
	  hr = E_FAIL;
	  basic_stringstream<WCHAR> strm;
	  strm << L"Item '" << Chk(bsTmp) << L"' is not found";
	  AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );		 
	}
   else
	{
	  rdwKey = it->second;
	  if( pIt ) *pIt = it;
	}

   bsTmp.Detach();
   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
HRESULT __fastcall CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::ExtractKeyFromIDispatch( IDispatch* pDsp, DWORD& rdwKey ) throw()
 {
   HRESULT hr = S_OK;

   
   CComPtr<IStCollItem> sp;
   hr = pDsp->QueryInterface( IID_IStCollItem, (void**)&sp );
   if( FAILED(hr) )
	 AtlReportError( *pCLSID, L"Cann't QI IStCollItem", *pIID, hr );		 
   else
	{
      long lSgn = 0;
	  sp->get_Sign( &lSgn );
	  if( lSgn != (long)static_cast<COLL*>(this)->GetUnknown() )
	   {
	     hr = E_FAIL;
		 basic_stringstream<WCHAR> strm;
		 CComBSTR bsTmp; sp->get_Name( &bsTmp );
	     strm << L"Item '" << Chk(bsTmp) << L"' is from another collection";
		 AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	   }
	  else
	   {
	     long lT;
	     hr = sp->get_Key( &lT );
		 rdwKey = lT;
	   }
	}
   
   return hr;
 }


template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::Update( VARIANT* KeyOrObj ) throw()
 {
   if( KeyOrObj == NULL ) return E_POINTER;   
   VARTYPE vt = (V_VT(KeyOrObj) & VT_TYPEMASK);

   HRESULT hr = S_OK;

   try {
	  DWORD dwKey;	  
	  IDispatch* pDsp;
	  CComPtr<IStCollItem> spSCI;
	  if( V_ISBYREF(KeyOrObj) == VT_BYREF )
	   {
		 switch( vt )
		  {	     
			case VT_DISPATCH:
			  hr = ExtractKeyFromIDispatch( *V_DISPATCHREF(KeyOrObj), dwKey );
			  pDsp = *V_DISPATCHREF(KeyOrObj);
			  break;
			default:
			  hr = E_INVALIDARG;
			  AtlReportError( *pCLSID, L"Bad type of KeyOrObj", *pIID, hr );
		  };
	   }
	  else
	   {
		 switch( vt )
		  {	     
			case VT_DISPATCH:
			   hr = ExtractKeyFromIDispatch( V_DISPATCH(KeyOrObj), dwKey );
			   pDsp = V_DISPATCH(KeyOrObj);
			   break;
			default:
			  hr = E_INVALIDARG;
			  AtlReportError( *pCLSID, L"Bad type of KeyOrObj", *pIID, hr );
		  };
	   }

	  if( SUCCEEDED(hr) &&
		  SUCCEEDED(hr = pDsp->QueryInterface(IID_IStCollItem, (void**)&spSCI))
		)
	   {
		 TFndByKey dta( dwKey );
		 IT_TMAP_DIR it = find_if( m_mapDir.begin(), m_mapDir.end(), dta );
		 if( it == m_mapDir.end() )
		  {
			hr = E_FAIL;
			AtlReportError( *pCLSID, L"Update: object is absent in directory", *pIID, hr );
		  }
		 else
		  {
			if( m_bCachedUpdates )
			 {
			   hr = E_FAIL;
			   AtlReportError( *pCLSID, L"Can't update in cached mode", *pIID, hr );
			 }
			else
			 { 			   
			   hr = StoreExisted( m_spStorage, spSCI, dwKey );

			   HRESULT hr2 = EndTransaction( m_spStorage, hr );
			   if( FAILED(hr2) && SUCCEEDED(hr) )
				 ReportStgError( hr2, L"Update", *pCLSID, *pIID ),
				 hr = hr2;

			   /*if( SUCCEEDED(hr) )
				{
				  CComBSTR bsTmp( it->first ); 					 
				  DWORD dwTmp = it->second;
				  m_mapDir.erase( it );
				  m_mapDir.insert( TMAP_DIR::value_type(bsTmp, dwTmp) );
				}*/		   
				
			 }//else
		  }//else
	   }//if
	 }//try
    catch( bad_alloc e )
	 {	  
	   hr = E_OUTOFMEMORY;
	 }
    catch( exception e )
	 {	  
	   //USES_CONVERSION;
	   hr = E_FAIL;
	   //AtlReportError( *pCLSID, A2COLE(e.what()), *pIID, hr );
	   ReportException( e, hr );
	 }

   return hr;
 }


template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get__NewEnum( LPUNKNOWN *ppVal ) throw()
 {   
   if( ppVal == NULL ) return E_POINTER;   
   
   HRESULT hr = S_OK;

   if( m_bCachedUpdates )
	{	  	     	  
	  _CachedEnum *pEnum = NULL;
	  pEnum = new (nothrow) CComObject<_CachedEnum>;
	  if( pEnum == NULL ) return E_POINTER;

	  pEnum->SetCachedUpdates( m_bCachedUpdates );
	  pEnum->SetDefaultCachedUpdates( m_bCachedUpdates );

	  size_t nCount = m_mapObj.size();
	  size_t uiSize = sizeof(IDispatch*) * nCount;
	  IDispatch** pVar = new (nothrow) IDispatch*[ nCount ];
	  if( pVar == NULL ) 
	   {
		 delete pEnum;
		 return E_POINTER;
	   }	  

	  // Set the VARIANTs that will hold the TYPE Interface pointers.	  
	  CIT_TMAP it;
	  IDispatch** pVi = pVar;
	  for( it = m_mapObj.begin(); it != m_mapObj.end(); it++, pVi++ )
        *pVi = reinterpret_cast<IDispatch*>( it->second.p );

	    //if( FAILED(hr = it->second.QueryInterface(pVi)) )
		  //break;	         	  	  
	  
	  TPF_CmpFunc pfn = GetCompareFunc_Objects();
	  if( pfn != NULL )
		qsort( pVar, nCount, sizeof(IDispatch*), pfn );
	  
	  IDispatch** pVEnd = pVar + nCount;
	  for( pVi = pVar; pVi != pVEnd; ++pVi )
	    if( FAILED(hr = reinterpret_cast<IStCollItem*>(*pVi)->QueryInterface( IID_IDispatch, (void**)pVi )) )
		  break;
	  
	  if( SUCCEEDED(hr) )
	   {
		 hr = pEnum->Init( pVar,
						   pVar + nCount, NULL,
						   static_cast<COLL*>(this)->GetUnknown() );
		 if( SUCCEEDED(hr) )     	 
		   hr = pEnum->QueryInterface( IID_IUnknown, (void**)ppVal );
	   }
	  

	  if( FAILED(hr) )
	   {
		 delete pEnum;	 
		 		
		 IDispatch **pKeep = pVar;
		 for( ;pVar < pVi; ++pVar )		  
		   if( *pVar ) (*pVar)->Release();
		  
		 delete[] pKeep;
	   }
	}
   else //not cached updates
	{
	  _NotCachedEnum *pEnum = NULL;
	  pEnum = new (nothrow) CComObject<_NotCachedEnum>;
	  if( pEnum == NULL ) return E_POINTER;

	  pEnum->SetCachedUpdates( m_bCachedUpdates );
	  pEnum->SetDefaultCachedUpdates( m_bCachedUpdates );

	  size_t nCount = m_mapDir.size();
	  size_t uiSize = sizeof(long) * nCount;
	  long* pVar = new (nothrow) long[ nCount ];
	  if( pVar == NULL ) 
	   {
		 delete pEnum;
		 return E_POINTER;
	   }	  
	  
	  // Set the VARIANTs that will hold the TYPE Interface pointers.
	  
	  CIT_TMAP_DIR it;
	  long* pVi = pVar;
	  for( it = m_mapDir.begin(); it != m_mapDir.end(); it++, pVi++ )
	    *pVi = it->second;
	   
	  
	  hr = pEnum->Init( pVar,
						pVar + nCount, m_spStorage,
						static_cast<COLL*>(this)->GetUnknown() );
	  if( SUCCEEDED(hr) )     	 
		hr = pEnum->QueryInterface( IID_IUnknown, (void**)ppVal );					   
	  
	  
	  if( FAILED(hr) )
	   {
		 delete pEnum;	 		
		 delete[] pVar;
	   }
	}

   if( FAILED(hr) ) *ppVal = NULL;
   
   return hr; 
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::ItemNth( long Number_, TYPE **ppGT ) throw()
 {
   HRESULT hr = S_OK;

   size_t Number = Number_;

   if( Number < 1 || Number > m_mapDir.size() )
	{
	  hr = E_FAIL;
	  basic_stringstream<WCHAR> strm;
	  strm << L"Number " << Number << L" is out of bounds";
	  AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	}
   else 
	{
	  IT_TMAP_DIR it = m_mapDir.begin();
	  advance( it, Number - 1 );
	  CComVariant vtTmp( (long)it->second, VT_I4 );
	  hr = Item( &vtTmp, ppGT );
	}

   return hr;
 };

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::KeyNth( /*[in]*/long Number, /*[out, retval]*/long* lKey ) throw()
 {
   HRESULT hr;   

   if( lKey == NULL ) hr = E_POINTER;
   if( Number < 1 || Number > m_mapDir.size() )
	{
	  hr = E_FAIL;
	  basic_stringstream<WCHAR> strm;
	  strm << L"Number " << Number << L" is out of bounds";
	  AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	}
   else 
	{
	  IT_TMAP_DIR it = m_mapDir.begin();
	  advance( it, Number - 1 );
	  *lKey = it->second;
	  hr = S_OK;
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::NameNth( /*[in]*/long Number, /*[out, retval]*/BSTR* bsName ) throw()
 {
   HRESULT hr;   

   if( bsName == NULL ) hr = E_POINTER;
   if( Number < 1 || Number > m_mapDir.size() )
	{
	  hr = E_FAIL;
	  basic_stringstream<WCHAR> strm;
	  strm << L"Number " << Number << L" is out of bounds";
	  AtlReportError( *pCLSID, strm.str().c_str(), *pIID, hr );
	}
   else 
	{
	  IT_TMAP_DIR it = m_mapDir.begin();
	  advance( it, Number - 1 );
	  hr = const_cast<CComBSTR*>(&it->first)->CopyTo( bsName );	  
	}

   return hr;
 }

template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::get_EnumSorting( StdSortSeqv* pSortSeq ) throw()
 {
   MGET_PROPERTY( pSortSeq, m_sssSeq );
   return S_OK;
 }


template< 
  class COLL,
  class TYPE,   
  class BASE, 
  const IID* pIID, 
  const CLSID* pCLSID,
  class PersistantImpl>
STDMETHODIMP CStColl<COLL,TYPE,BASE,pIID,pCLSID,PersistantImpl>::put_EnumSorting( StdSortSeqv SortSeq ) throw()
 {
   MPUT_PROPERTY_NM( m_sssSeq, SortSeq );
   return S_OK;
 }



//#pragma warning( pop )

#endif

