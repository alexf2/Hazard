// CollSF.h : Declaration of the CCollSF

#ifndef __COLLSF_H_
#define __COLLSF_H_

#include "resource.h"       // main symbols

#include "CollOnMap.h"
#include "PassErr.h"
#include "SafetyPrecaution.h"

#include <algorithm>
#include <list>
#include <sstream>
#include <iomanip>
using namespace std;


typedef  CComCollectionOnMap< ISafetyPrecaution, 
	  long,
	  IDispatchImpl<ICollSF, &IID_ICollSF, &LIBID_GERTNETLib> > CCL4;

inline int IsDrt_CCollSF( CCL4::TMAP::value_type& rVal )
 {
   if( rVal.second == NULL ) return false;
   CComQIPtr<IPersistStreamInit> spSI( rVal.second );
   if( !spSI ) return false;	
   else return spSI->IsDirty() == S_OK ? true:false;
 }

class CCollSF;
typedef CIClonableImpl_CollOnMap<CCollSF, CCL4::TMAP, ISafetyPrecaution> TCloneImpl4;

/////////////////////////////////////////////////////////////////////////////
// CCollSF
class ATL_NO_VTABLE CCollSF : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCollSF, &CLSID_CollSF>,
	public ISupportErrorInfo,
	//public IDispatchImpl<ICollSF, &IID_ICollSF, &LIBID_GERTNETLib>
	public CCL4,
	public IPersistStorage,	
	public TCloneImpl4
	//public IClonable
{
public:
	CCollSF()
	{
	  m_csfSorting = CSFS_None;
	}

	~CCollSF()
	 {
	   //int i=1;
	 }

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid )
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   static const IID* arr[] = 
	   {
		   &IID_ICollSF,
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

DECLARE_REGISTRY_RESOURCEID(IDR_COLLSF)
DECLARE_NOT_AGGREGATABLE(CCollSF)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCollSF)
	COM_INTERFACE_ENTRY(ICollSF)
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

		
	   ATLTRACE2(atlTraceCOM, 0, _T("ICollSF::LookNextIndex\n"));	   

       if( plIndex == NULL ) return E_POINTER;
	   *plIndex = m_dwMaxKey + 1;
	   return S_OK;
	 }
	

public:
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   HRESULT hr = TCloneImpl4::Clone( ppUnk );
	   
	   if( SUCCEEDED(hr) )	    	    
		 m_pObjTmp->m_dwMaxKey = m_dwMaxKey,
		 m_pObjTmp->m_bRequiresSave = false,
		 m_pObjTmp->m_csfSorting = m_csfSorting,

		 m_pObjTmp->m_ocProfit = m_ocProfit,
		 m_pObjTmp->m_dKe = m_dKe,
	     m_pObjTmp->m_ocCost = m_ocCost,
		 m_pObjTmp->m_dDeltaQ = m_dDeltaQ; 
	   
		
       return hr;		
	 }

public:
	STDMETHOD(CheckCompatible)(ICollSF** pSFNone,/*[out,retval]*/VARIANT_BOOL* bResult);
	STDMETHOD(get_DeltaQ)(/*[out, retval]*/ double *pVal);
	STDMETHOD(put_DeltaQ)(/*[in]*/ double newVal);
	STDMETHOD(get_Cost)(/*[out, retval]*/ CURRENCY *pVal);
	STDMETHOD(put_Cost)(/*[in]*/ CURRENCY newVal);
	STDMETHOD(get_Ke)(/*[out, retval]*/ double *pVal);
	STDMETHOD(put_Ke)(/*[in]*/ double newVal);
	STDMETHOD(get_Profit)(/*[out, retval]*/ CURRENCY *pVal);
	STDMETHOD(put_Profit)(/*[in]*/ CURRENCY newVal);
	STDMETHOD(get_IsDirty)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_Sorting)(/*[out, retval]*/ CollSFSorting *pVal);
	STDMETHOD(put_Sorting)(/*[in]*/ CollSFSorting newVal);
	STDMETHOD(InvalidateDelta)();
	STDMETHOD(get_IsValidDelta)(/*[out, retval]*/ VARIANT_BOOL *pVal);

  CComObject<CCollSF>* m_pObjTmp; //for CIClonableImpl_CollOnMap

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
	   TMAP::iterator itRes = find_if( m_coll.begin(), m_coll.end(), IsDrt_CCollSF );

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
		m_csfSorting = CSFS_None;

		CY cy; cy.int64 = 0;   
		m_ocProfit = cy;
		m_dKe = 0;

		m_ocCost = cy;
		m_dDeltaQ = 0;

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
	     PassError( L"IStorage.EnumElements", hr, GetObjectCLSID(), IID_ICollSF  );   
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

			CComObject<CSafetyPrecaution>* pLE;
			hr = CComObject<CSafetyPrecaution>::CreateInstance( &pLE );
			if( FAILED(hr) )
			 {
               PassError( L"CComObject<CSafetyPrecaution>::CreateInstance", hr, GetObjectCLSID(), IID_ICollSF  );   			   
	           return hr;
			 }

			CComPtr<IPersistStreamInit> spPSI;
			hr = pLE->_InternalQueryInterface( IID_IPersistStreamInit, (void**)&spPSI );
			if( FAILED(hr) )
			 {
               PassError( L"_InternalQueryInterface", hr, GetObjectCLSID(), IID_ICollSF  );   			   
			   delete pLE;
	           return hr;
			 }

			CComPtr<IStream> spInfoStream;
			hr = pStorage->OpenStream( Chk(bsName), NULL, 
	          STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spInfoStream );
			if( FAILED(hr) ) 
			 {			   
	           PassError( L"IStorage.OpenStream", hr, GetObjectCLSID(), IID_ICollSF  );   			   
			   spPSI.Release(); 
			   //delete pLE;
	           return hr;
			 }

			hr = spPSI->Load( spInfoStream );
			if( FAILED(hr) ) 
			 {			   
	           PassError( L"IStreamInitNew.Load", hr, GetObjectCLSID(), IID_ICollSF  );   			   
			   spPSI.Release(); 
			   //delete pLE;
	           return hr;
			 }

			CComQIPtr<ISafetyPrecaution> spLEnum( spPSI );
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
	   TItemManip( ISafetyPrecaution* p, DWORD dw )
		{
		  m_p = p; m_dw = dw;
		}
	   ISafetyPrecaution* m_p;
	   DWORD m_dw;
	 };

	struct TFndItDta
	 {
       TFndItDta( ISafetyPrecaution* p )
		{
		  m_p = p;
		}

	   bool operator()( const TItemManip& rI )
		{
          return rI.m_p == m_p;
		}

	   ISafetyPrecaution* m_p;
	 };

	struct TFndItDta_TAMP
	 {
       TFndItDta_TAMP( ISafetyPrecaution* p )
		{
		  m_p = p;
		}

	   bool operator()( const TMAP::value_type& rI )
		{
          return rI.second == m_p;
		}

	   ISafetyPrecaution* m_p;
	 };
	

	STDMETHOD(Save)(IStorage* pStorage, BOOL fSameAsLoad)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	  ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::Save\n"));

	  if( !pStorage ) return E_POINTER;
      	  

      list<TItemManip> lstAdd, lstUpdate;
	  list<DWORD> lstRemove;

	  CComPtr<IEnumSTATSTG> spEnum;
	  HRESULT hr = pStorage->EnumElements( 0, NULL, 0, &spEnum );
	  if( FAILED(hr) ) 
	   {
	     PassError( L"IStorage.EnumElements", hr, GetObjectCLSID(), IID_ICollSF  );   
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
		 if( it == lstUpdate.end() )// && CComQIPtr<IPersistStreamInit>(itSta->second)->IsDirty() == S_OK )
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
			PassError( L"IStorage.DestroyElement", hr, GetObjectCLSID(), IID_ICollSF  );   
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
			PassError( L"IStorage.OpenStream", hr, GetObjectCLSID(), IID_ICollSF  );   
			return hr;
		  }

		 CComQIPtr<IPersistStreamInit> spInitNew( itS->m_p );
		 if( !spInitNew )
		  {
            PassError( L"QI IStreamInitNew", E_FAIL, GetObjectCLSID(), IID_ICollSF  );   
			return E_FAIL;
		  }

		 ULARGE_INTEGER ulInt; ulInt.QuadPart = 0;
         hr = spStrm->SetSize( ulInt ); 
		 if( FAILED(hr) )
		  {
			PassError( L"IStream.SetSize", hr, GetObjectCLSID(), IID_ICollSF  );   
			return hr;
		  }

		 LARGE_INTEGER lInt; lInt.QuadPart = 0;
		 spStrm->Seek( lInt, STREAM_SEEK_SET, NULL );

		 hr = spInitNew->Save( spStrm, TRUE );
		 if( FAILED(hr) )
		  {
			PassError( L"IPersistStreamInit.Save", hr, GetObjectCLSID(), IID_ICollSF  );   
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
			PassError( L"IStorage.CreateStream", hr, GetObjectCLSID(), IID_ICollSF );   
			return hr;
		  }

		 CComQIPtr<IPersistStreamInit> spInitNew( itS->m_p );
		 if( !spInitNew )
		  {
            PassError( L"QI IStreamInitNew", E_FAIL, GetObjectCLSID(), IID_ICollSF );   
			return E_FAIL;
		  }         

		 hr = spInitNew->Save( spStrm, TRUE );
		 if( FAILED(hr) )
		  {
			PassError( L"IPersistStreamInit.Save", hr, GetObjectCLSID(), IID_ICollSF  );   
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

   
	STDMETHOD(get__NewEnum)( LPUNKNOWN *ppVal )
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


		HRESULT hr = S_OK;

		if( ppVal == NULL )
			return E_POINTER;
		*ppVal = NULL;

		_CComCollEnum *pEnum = NULL;
		pEnum = new (nothrow) CComObject<_CComCollEnum>();
		if( pEnum == NULL ) return E_POINTER;
    
		// allocate an initialize a vector of VARIANT objects
		int nCount = m_coll.size();
		DWORD uiSize = sizeof(VARIANT) * (DWORD)nCount;
		VARIANT* pVar = (VARIANT*)new (nothrow) VARIANT[ nCount ];
		if( pVar == NULL ) 
		 {
		   delete pEnum;
		   return E_POINTER;
		 }
		memset( pVar, '\0', uiSize );

		// Set the VARIANTs that will hold the TYPE Interface pointers.
		int i = 0;
		TMAP::const_iterator it;
		VARIANT* pVi = pVar;
		for( it = m_coll.begin(); it != m_coll.end(); it++, pVi++ )
		 {			    
			V_VT(pVi) = VT_DISPATCH;
			V_DISPATCH(pVi) = NULL;
			
			if( it->second != NULL )	
			  if( FAILED(hr = it->second->QueryInterface( IID_IDispatch, (void**)&V_DISPATCH(pVi) )) ) 
				 break;
						 
			i++;
		 }		

		if( SUCCEEDED(hr) )
		 {
		   MakeSorting( pVar, nCount );

		   hr = pEnum->Init( &pVar[0],
							 &pVar[nCount],
							 static_cast<IDispatch*>(this)/*NULL*/, AtlFlagTakeOwnership );
		   if( SUCCEEDED(hr) )     	 
			 hr = pEnum->QueryInterface( IID_IUnknown, (void**)ppVal );					   
		 }

		if( FAILED(hr) )
		 {
		   delete pEnum;	 
		   
		   --i;
		   while( i >= 0 )
			{
			   VariantClear( pVar + i );
			   i--;
			}
		   delete[] pVar;
		 }
		
		return hr; 
	 }

	STDMETHOD(Remove)( CCL4::ITEMMAP_B index )
	 {
	   HRESULT hr = CCL4::Remove( index );	   
	   if( SUCCEEDED(hr) )
		{
		  TMAP::iterator it( m_coll.begin() );
          for( ; it != m_coll.end(); ++it )
		   {
             CComObject<CSafetyPrecaution>& rSP = *static_cast<CComObject<CSafetyPrecaution>*>( it->second );
			 vector<long>::iterator itV = find( rSP.m_vecCollNC.begin(), rSP.m_vecCollNC.end(), index );
			 if( itV != rSP.m_vecCollNC.end() )
			   rSP.m_vecCollNC.erase( itV );
		   }
		}
	   return hr;
	 }

  
protected:
  CollSFSorting m_csfSorting;

  COleCurrency m_ocProfit, m_ocCost;
  double m_dKe, m_dDeltaQ;

  void MakeSorting( VARIANT* pVar, unsigned long uiSize )
   {
     typedef int (__cdecl *pfnCmpTyp)( const void *p1, const void *p2 );
	 pfnCmpTyp pfnCmp;

	 switch( m_csfSorting )
	  {
	     case CSFS_Price:
		    pfnCmp = CmpPrices; break;
         case CSFS_DeltaQ:
		    pfnCmp = CmpDeltaQ; break;
		 case CSFS_PriceDescending:
		    pfnCmp = CmpPricesDesc; break;
		 case CSFS_DeltaQDescending:
		    pfnCmp = CmpDeltaQDesc; break;
		 case CSFS_KDescending:
		    pfnCmp = CmpKDesc; break;
		 default:
		   return;
	  };	 
	 
	 qsort( pVar, uiSize, sizeof(VARIANT), pfnCmp );	 
   }
  static int __cdecl CmpPrices( const void *p1, const void *p2 )
   {
     CComObject<CSafetyPrecaution> *pF1, *pF2; 
	 VARIANT& v1 = *(VARIANT*)p1;
	 VARIANT& v2 = *(VARIANT*)p2;
	 //if( v1.vt == VT_DISPATCH )
	   pF1 = static_cast<CComObject<CSafetyPrecaution>*>( v1.pdispVal ),
	   pF2 = static_cast<CComObject<CSafetyPrecaution>*>( v2.pdispVal );
	 /*else	   
	   pF1 = static_cast<CComObject<CSafetyPrecaution>*>( v1.punkVal ),
	   pF2 = static_cast<CComObject<CSafetyPrecaution>*>( v2.punkVal );*/

     if( pF1->m_ocCost < pF2->m_ocCost ) return -1;
	 else if( pF1->m_ocCost == pF2->m_ocCost ) return 0;
	 else return 1;
   }
  static int __cdecl CmpDeltaQ( const void *p1, const void *p2 )
   {
     CComObject<CSafetyPrecaution> *pF1, *pF2; 
	 VARIANT& v1 = *(VARIANT*)p1;
	 VARIANT& v2 = *(VARIANT*)p2;
	 //if( v1.vt == VT_DISPATCH )
	   pF1 = static_cast<CComObject<CSafetyPrecaution>*>( v1.pdispVal ),
	   pF2 = static_cast<CComObject<CSafetyPrecaution>*>( v2.pdispVal );
	 /*else	   
	   pF1 = static_cast<CComObject<CSafetyPrecaution>*>( v1.punkVal ),
	   pF2 = static_cast<CComObject<CSafetyPrecaution>*>( v2.punkVal );*/

     if( pF1->m_dDeltaQ < pF2->m_dDeltaQ ) return -1;
	 else if( pF1->m_dDeltaQ == pF2->m_dDeltaQ ) return 0;
	 else return 1;
   }

  static int __cdecl CmpPricesDesc( const void *p1, const void *p2 )
   {
     CComObject<CSafetyPrecaution> *pF1, *pF2; 
	 VARIANT& v1 = *(VARIANT*)p1;
	 VARIANT& v2 = *(VARIANT*)p2;
	
	   pF1 = static_cast<CComObject<CSafetyPrecaution>*>( v1.pdispVal ),
	   pF2 = static_cast<CComObject<CSafetyPrecaution>*>( v2.pdispVal );
	

     if( pF1->m_ocCost > pF2->m_ocCost ) return -1;
	 else if( pF1->m_ocCost == pF2->m_ocCost ) return 0;
	 else return 1;
   }

  static int __cdecl CmpDeltaQDesc( const void *p1, const void *p2 )
   {
     CComObject<CSafetyPrecaution> *pF1, *pF2; 
	 VARIANT& v1 = *(VARIANT*)p1;
	 VARIANT& v2 = *(VARIANT*)p2;
	 
	 pF1 = static_cast<CComObject<CSafetyPrecaution>*>( v1.pdispVal ),
	 pF2 = static_cast<CComObject<CSafetyPrecaution>*>( v2.pdispVal );
	 

     if( pF1->m_dDeltaQ > pF2->m_dDeltaQ ) return -1;
	 else if( pF1->m_dDeltaQ == pF2->m_dDeltaQ ) return 0;
	 else return 1;
   }

  static int __cdecl CmpKDesc( const void *p1, const void *p2 )
   {
     CComObject<CSafetyPrecaution> *pF1, *pF2; 
	 VARIANT& v1 = *(VARIANT*)p1;
	 VARIANT& v2 = *(VARIANT*)p2;
	 
	 pF1 = static_cast<CComObject<CSafetyPrecaution>*>( v1.pdispVal ),
	 pF2 = static_cast<CComObject<CSafetyPrecaution>*>( v2.pdispVal );
	 

     if( pF1->m_dKe > pF2->m_dKe ) return -1;
	 else if( pF1->m_dKe == pF2->m_dKe ) return 0;
	 else return 1;
   }
  
};

#endif //__COLLSF_H_
