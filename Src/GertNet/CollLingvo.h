// CollLingvo.h : Declaration of the CCollLingvo

#ifndef __COLLLINGVO_H_
#define __COLLLINGVO_H_

#include "resource.h"       // main symbols
//#include <OtlComCollectionSTL.h>
#include "CollOnMap.h"
#include "PassErr.h"
#include "LingvoEnum.h"

#include "CollFunctors.h"

#include <algorithm>
#include <list>
#include <sstream>
#include <iomanip>
using namespace std;


class CCollLingvo;

typedef  CComCollectionOnMap< ILingvoEnum, 
	  long,
	  IDispatchImpl<ICollLingvo, &IID_ICollLingvo, &LIBID_GERTNETLib> > CCL;

inline int IsDrt_CCollLingvo( CCL::TMAP::value_type& rVal )
 {
   if( rVal.second == NULL ) return false;
   CComQIPtr<IPersistStreamInit> spSI( rVal.second );
   if( !spSI ) return false;	
   else return spSI->IsDirty() == S_OK ? true:false;
 }


typedef CIClonableImpl_CollOnMap<CCollLingvo, CCL::TMAP, ILingvoEnum> TCloneImpl;


/////////////////////////////////////////////////////////////////////////////
// CCollLingvo
class ATL_NO_VTABLE CCollLingvo : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCollLingvo, &CLSID_CollLingvo>,
	//public CComCollectionOnSTLMap<ILingvoEnum, IDispatchImpl<ICollLingvo, &IID_ICollLingvo, &LIBID_GERTNETLib> >
	public CCL,
	public IPersistStorage,
	public ISupportErrorInfo,	
	public TCloneImpl
	//public IClonable
{

friend struct CApplyDirty<CCollLingvo>;

public:
	CCollLingvo():m_pObjTmp( NULL )
	{	  	  
	}

	STDMETHOD(get_IsDirty)(/*[out, retval]*/ VARIANT_BOOL *pVal)
	 {
       if( !pVal ) return E_POINTER;

	   *pVal = (IsDirty() == S_OK ? VARIANT_TRUE:VARIANT_FALSE);
	   return S_OK;
	 }
	
	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid )
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   static const IID* arr[] = 
	   {
		   &IID_ICollLingvo,
		   &IID_IPersistStorage,
		   &IID_IPersistStreamInit,		
		   &IID_IClonable,
		   &IID_ILongKey
	   };
	   for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	   {
		   if (InlineIsEqualGUID(*arr[i],riid))
			   return S_OK;
	   }
	   return S_FALSE;
	 }

//IPersistStorage
	STDMETHOD(GetClassID)( CLSID *pClassID )
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::GetClassID\n"));
	   *pClassID = GetObjectCLSID();
	   return S_OK;
	 }

	
	STDMETHOD(IsDirty)(void)
	 {	   
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::IsDirty\n"));		
	   
	   if( m_bRequiresSave ) return S_OK;
   
	   bool bFlFnd = false;   	   
	   TMAP::iterator itRes = find_if( m_coll.begin(), m_coll.end(), IsDrt_CCollLingvo );

	   return itRes != m_coll.end() ? S_OK:S_FALSE;
	 }
	
	STDMETHOD(InitNew)(IStorage*)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::InitNew\n"));
		RemoveAll();
		m_dwMaxKey = 0;
		m_bRequiresSave = true;

		return S_OK;
	}
	STDMETHOD(Load)(IStorage* pStorage)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	  ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::Load\n"));

	  if( !pStorage ) return E_POINTER;

	  RemoveAll();

	  CComPtr<IEnumSTATSTG> spEnum;
	  HRESULT hr = pStorage->EnumElements( 0, NULL, 0, &spEnum );
	  if( FAILED(hr) ) 
	   {
	     PassError( L"IStorage.EnumElements", hr, GetObjectCLSID(), IID_ICollLingvo  );   
	     return hr;
	   }

	  m_dwMaxKey = 0;
	  spEnum->Reset();
	  STATSTG stg; memset( &stg, 0, sizeof(STATSTG) );
	  while( spEnum->Next(1, &stg, NULL) == S_OK )
	   {
	     CComBSTR bsName( stg.pwcsName );
		 if( stg.pwcsName ) CoTaskMemFree( stg.pwcsName ), stg.pwcsName = NULL;

         if( (stg.type & STGTY_STREAM) == STGTY_STREAM ) 
		  {
            DWORD dwNum = wcstoul( Chk(bsName), NULL, 10 );

			CComObject<CLingvoEnum>* pLE;
			hr = CComObject<CLingvoEnum>::CreateInstance( &pLE );
			if( FAILED(hr) )
			 {
               PassError( L"CComObject<CLingvoEnum>::CreateInstance", hr, GetObjectCLSID(), IID_ICollLingvo  );   			   
	           return hr;
			 }

			CComPtr<IPersistStreamInit> spPSI;
			hr = pLE->_InternalQueryInterface( IID_IPersistStreamInit, (void**)&spPSI );
			if( FAILED(hr) )
			 {
               PassError( L"_InternalQueryInterface", hr, GetObjectCLSID(), IID_ICollLingvo  );   			   
			   delete pLE;
	           return hr;
			 }

			CComPtr<IStream> spInfoStream;
			hr = pStorage->OpenStream( Chk(bsName), NULL, 
	          STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spInfoStream );
			if( FAILED(hr) ) 
			 {			   
	           PassError( L"IStorage.OpenStream", hr, GetObjectCLSID(), IID_ICollLingvo  );   			   
			   spPSI.Release(); 
			   //delete pLE;
	           return hr;
			 }

			hr = spPSI->Load( spInfoStream );
			if( FAILED(hr) ) 
			 {			   
	           PassError( L"IStreamInitNew.Load", hr, GetObjectCLSID(), IID_ICollLingvo  );   			   
			   spPSI.Release(); 
			   //delete pLE;
	           return hr;
			 }

			CComQIPtr<ILingvoEnum> spLEnum( spPSI );
			spPSI.Release(); 

			m_coll.insert( TMAP::value_type(dwNum, spLEnum) );
			m_dwMaxKey = max( m_dwMaxKey, dwNum );
			spLEnum.Detach();
		  }

         
		 memset( &stg, 0, sizeof(STATSTG) );
	   }

	  m_bRequiresSave = false;

	  return S_OK;
	}

    struct TItemManip
	 {
	   TItemManip( ILingvoEnum* p, DWORD dw )
		{
		  m_p = p; m_dw = dw;
		}
	   ILingvoEnum* m_p;
	   DWORD m_dw;
	 };

	struct TFndItDta
	 {
       TFndItDta( ILingvoEnum* p )
		{
		  m_p = p;
		}

	   bool operator()( const TItemManip& rI )
		{
          return rI.m_p == m_p;
		}

	   ILingvoEnum* m_p;
	 };

	struct TFndItDta_TAMP
	 {
       TFndItDta_TAMP( ILingvoEnum* p )
		{
		  m_p = p;
		}

	   bool operator()( const TMAP::value_type& rI )
		{
          return rI.second == m_p;
		}

	   ILingvoEnum* m_p;
	 };
	

	STDMETHOD(Save)(IStorage* pStorage, BOOL fSameAsLoad)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	  ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::Save\n"));

	  if( !pStorage ) return E_POINTER;
      
	  /*TMAP::iterator itSta( m_coll.begin() );
	  TMAP::iterator itEnd( m_coll.end() );
	  for( ; itSta != itEnd; ++itSta )
	   {
         CComQIPtr<IPersistStreamInit> spSI( itSta->second );
		 if( !spSI || spSI->IsDirty() == S_FALSE ) continue;         
	   }*/				

      list<TItemManip> lstAdd, lstUpdate;
	  list<DWORD> lstRemove;

	  CComPtr<IEnumSTATSTG> spEnum;
	  HRESULT hr = pStorage->EnumElements( 0, NULL, 0, &spEnum );
	  if( FAILED(hr) ) 
	   {
	     PassError( L"IStorage.EnumElements", hr, GetObjectCLSID(), IID_ICollLingvo  );   
	     return hr;
	   }

	  spEnum->Reset();
	  STATSTG stg; memset( &stg, 0, sizeof(STATSTG) );
	  while( spEnum->Next(1, &stg, NULL) == S_OK )
	   {
	     CComBSTR bsName( stg.pwcsName );
		 if( stg.pwcsName ) CoTaskMemFree( stg.pwcsName ), stg.pwcsName = NULL;

         if( (stg.type & STGTY_STREAM) == STGTY_STREAM ) 
		  {
            DWORD dwNum = wcstoul( Chk(bsName), NULL, 10 );

		    TMAP::iterator it = m_coll.find( dwNum );

			if( it == m_coll.end() ) lstRemove.push_back( dwNum );
			else 
			 {
               CComQIPtr<IPersistStreamInit> spSI( it->second );
               if( spSI.p && spSI->IsDirty() == S_OK ) lstUpdate.push_back( TItemManip(it->second, it->first) );
			 }			
		  }
         
		 memset( &stg, 0, sizeof(STATSTG) );
	   }

	  spEnum = NULL;

	  TMAP::iterator itSta( m_coll.begin() );
	  TMAP::iterator itEnd( m_coll.end() );
	  for( ; itSta != itEnd; ++itSta )
	   {
	     TFndItDta dta( itSta->second );
	     list<TItemManip>::iterator it = find_if( lstUpdate.begin(), lstUpdate.end(), dta );
		 if( it == lstUpdate.end() )//&& CComQIPtr<IPersistStreamInit>(itSta->second)->IsDirty() == S_OK )
		   lstAdd.push_back( TItemManip(itSta->second, itSta->first) );		
	   }


	  list<DWORD>::iterator it1( lstRemove.begin() );
	  list<DWORD>::iterator it2( lstRemove.end() );
	  for( ; it1 != it2; ++it1 )
       {
		 wchar_t wcBuff[ 64 ];
	     swprintf( wcBuff, L"%lu", *it1 );
         hr = pStorage->DestroyElement( wcBuff );
		 if( FAILED(hr) ) 
		  {
			PassError( L"IStorage.DestroyElement", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			return hr;
		  }
	   }

	  list<TItemManip>::iterator itS( lstUpdate.begin() );
	  list<TItemManip>::iterator itE( lstUpdate.end() );
	  for( ; itS != itE; ++itS )
	   {
         CComPtr<IStream> spStrm;
		 wchar_t wcBuff[ 64 ];
	     swprintf( wcBuff, L"%lu", itS->m_dw );
		 hr = pStorage->OpenStream( wcBuff, NULL,
	       STGM_WRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spStrm );
		 if( FAILED(hr) )
		  {
			PassError( L"IStorage.OpenStream", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			return hr;
		  }

		 CComQIPtr<IPersistStreamInit> spInitNew( itS->m_p );
		 if( !spInitNew )
		  {
            PassError( L"QI IStreamInitNew", E_FAIL, GetObjectCLSID(), IID_ICollLingvo  );   
			return E_FAIL;
		  }

		 ULARGE_INTEGER ulInt; ulInt.QuadPart = 0;
         hr = spStrm->SetSize( ulInt ); 
		 if( FAILED(hr) )
		  {
			PassError( L"IStream.SetSize", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			return hr;
		  }

		 LARGE_INTEGER lInt; lInt.QuadPart = 0;
		 spStrm->Seek( lInt, STREAM_SEEK_SET, NULL );

		 hr = spInitNew->Save( spStrm, TRUE );
		 if( FAILED(hr) )
		  {
			PassError( L"IPersistStreamInit.Save", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			return hr;
		  }
	   }
	   
	   //lstAdd, lstUpdate;
	  itS = lstAdd.begin();
	  itE = lstAdd.end();
	  for( ; itS != itE; ++itS )
	   {
         CComPtr<IStream> spStrm;
		 wchar_t wcBuff[ 64 ];
	     swprintf( wcBuff, L"%lu", itS->m_dw );
		 hr = pStorage->CreateStream( wcBuff,
	       STGM_CREATE|STGM_WRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spStrm );
		 if( FAILED(hr) )
		  {
			PassError( L"IStorage.CreateStream", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			return hr;
		  }

		 CComQIPtr<IPersistStreamInit> spInitNew( itS->m_p );
		 if( !spInitNew )
		  {
            PassError( L"QI IStreamInitNew", E_FAIL, GetObjectCLSID(), IID_ICollLingvo  );   
			return E_FAIL;
		  }         

		 hr = spInitNew->Save( spStrm, TRUE );
		 if( FAILED(hr) )
		  {
			PassError( L"IPersistStreamInit.Save", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			return hr;
		  }
	   }

	  m_bRequiresSave = false;

	  return S_OK;
	}
	STDMETHOD(SaveCompleted)(IStorage* /* pStorage */)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::SaveCompleted\n"));
		return S_OK;
	}
	STDMETHOD(HandsOffStorage)(void)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::HandsOffStorage\n"));
		return S_OK;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_COLLLINGVO)
DECLARE_NOT_AGGREGATABLE(CCollLingvo)

BEGIN_COM_MAP(CCollLingvo)
	COM_INTERFACE_ENTRY(ICollLingvo)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)
	COM_INTERFACE_ENTRY(IClonable)
END_COM_MAP()

// ICollLingvo
public:
	STDMETHOD(LookNextIndex)(/*[out, retval]*/long* plIndex)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		
	   ATLTRACE2(atlTraceCOM, 0, _T("ICollLingvo::LookNextIndex\n"));	   

       if( plIndex == NULL ) return E_POINTER;
	   *plIndex = m_dwMaxKey + 1;
	   return S_OK;
	 }

	STDMETHOD(UpdateFrom)(/*[in]*/ICollLingvo* pIColl)
	 {
	   if( pIColl == NULL ) return E_POINTER;

	   list<long> lAdd, lUpd, lRemov;
	   CComObject<CCollLingvo>* pColl = static_cast<CComObject<CCollLingvo>*>( pIColl );
	   TMAP::iterator it( pColl->m_coll.begin() );
	   for( ; it != pColl->m_coll.end(); ++it )
		{
          CComObject<CLingvoEnum>* ple = static_cast<CComObject<CLingvoEnum>*>( it->second );
		  TMAP::iterator itFnd = m_coll.find( it->first );
		  if( itFnd != m_coll.end() )
		   {
		     if( ple->IsDirty() == S_OK ) lUpd.push_back( it->first );
		   }
		  else lAdd.push_back( it->first );
		}

	   for( it = m_coll.begin(); it != m_coll.end(); ++it )
		{
          TMAP::iterator itFnd = pColl->m_coll.find( it->first );
		  if( itFnd == pColl->m_coll.end() )
		    lRemov.push_back( it->first );
		}

	   list<long>::iterator itOp( lRemov.begin() );
	   for( ; itOp != lRemov.end(); ++itOp )
		{
		  HRESULT hr = Remove( *itOp );
		  if( FAILED(hr) )
		   {
		     PassError( L"UpdateFrom", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			 return hr;
		   }
		}

	   for( itOp = lAdd.begin(); itOp != lAdd.end(); ++itOp )
		{
		  HRESULT hr = Add( pColl->m_coll.find(*itOp)->second, *itOp, NULL );
		  if( FAILED(hr) )
		   {
		     PassError( L"UpdateFrom", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			 return hr;
		   }
		}

	   for( itOp = lUpd.begin(); itOp != lUpd.end(); ++itOp )
		{
          CComObject<CLingvoEnum>* pleDst = 
		    static_cast<CComObject<CLingvoEnum>*>( m_coll.find(*itOp)->second );

		  //CComObject<CLingvoEnum>* pleSrc = 
		    //static_cast<CComObject<CLingvoEnum>*>( pColl->m_coll.find(*itOp)->second );

		  HRESULT hr = pleDst->UpdateFrom( pColl->m_coll.find(*itOp)->second );
		  if( FAILED(hr) )
		   {
		     PassError( L"UpdateFrom", hr, GetObjectCLSID(), IID_ICollLingvo  );   
			 return hr;
		   }
		}

	   return S_OK;
	 }
	

public:
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk)
	 {
	   #ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	   

	   /*if( ppUnk == NULL ) return E_POINTER;

		CComObject<CCollLingvo>* pObj;
		HRESULT hr = CComObject<CCollLingvo>::CreateInstance( &pObj );
		if( FAILED(hr) )
		 {
		   PassError( L"CCollLingvo::Copy: CComObject<CCollLingvo>::CreateInstance", hr, GetObjectCLSID(), IID_ICollLingvo );
		   return hr;
		 }

		CComPtr<IUnknown> spTmp;
		hr = pObj->_InternalQueryInterface( IID_IUnknown, (void**)&spTmp );
		if( FAILED(hr) )
		 {
		   delete pObj;
		   PassError( L"CCollLingvo::Copy: CComObject<CCollLingvo>::_InternalQueryInterface", hr, GetObjectCLSID(), IID_ICollLingvo );
		   return hr;
		 }

		TMAP::iterator it1( m_coll.begin() );
		TMAP::iterator it2( m_coll.end() );
		for( ; it1 != it2; ++it1 )
		 {
		   CComQIPtr<IClonable> spClon( it1->second );
		   if( !spClon )
			{
			  PassError( L"CCollLingvo::Copy: CComQIPtr<IClonable>", E_FAIL, GetObjectCLSID(), IID_ICollLingvo );
		      return E_FAIL;
			}
		   CComPtr<IUnknown> spCopy;
		   hr = spClon->Clone( &spCopy );
		   if( FAILED(hr) )
			{			 
			  PassError( L"CCollLingvo::Copy: spClon->Clone", hr, GetObjectCLSID(), IID_ICollLingvo );
			  return hr;
			}
		   
		   CComQIPtr<ILingvoEnum> spFac( spCopy.p );
		   spFac.p->AddRef();
		   pObj->m_coll.insert( TMAP::value_type(it1->first, spFac.p) );
		 }*/

		HRESULT hr = TCloneImpl::Clone( ppUnk );

	   if( SUCCEEDED(hr) )	    
		 m_pObjTmp->m_dwMaxKey = m_dwMaxKey,
		 m_pObjTmp->m_bRequiresSave = false;
		
		
		return hr;
	 }

public:	
	
  CComObject<CCollLingvo>* m_pObjTmp; //for CIClonableImpl_CollOnMap

private:
  
};

#endif //__COLLLINGVO_H_
