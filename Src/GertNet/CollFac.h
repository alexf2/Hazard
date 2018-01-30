// CollFac.h : Declaration of the CCollFac

#ifndef __COLLFAC_H_
#define __COLLFAC_H_

#include "resource.h"       // main symbols
//#include <OtlComCollectionSTL.h>
#include "CollOnMap.h"
#include "PassErr.h"
#include "Factor.h"

#include "CollFunctors.h"

#include <comutil.h>
#include <comdef.h>
#include <algorithm>
#include <list>
#include <sstream>
#include <iomanip>
using namespace std;


/////////////////////////////////////////////////////////////////////////////
// CCollFac


typedef  CComCollectionOnMap_BSTR< IFactor, 	  
	  IDispatchImpl<ICollFac, &IID_ICollFac, &LIBID_GERTNETLib> > CCL2;

inline int IsDrt_CCollFac( CCL2::TMAP::value_type& rVal )
 {
   if( rVal.second == NULL ) return false;
   CComQIPtr<IPersistStreamInit> spSI( rVal.second );
   if( !spSI ) return false;	
   else return spSI->IsDirty() == S_OK ? true:false;
 }

class CCollFac;
typedef CIClonableImpl_CollOnMap<CCollFac, CCL2::TMAP, IFactor> TCloneImpl2;


class ATL_NO_VTABLE CCollFac : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCollFac, &CLSID_CollFac>,
	public ISupportErrorInfo,
	public CCL2,
	public IPersistStorage,    
	public TCloneImpl2
	//public IClonable

{

friend struct CApplyDirty<CCollFac>;

public:
	CCollFac()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_COLLFAC)
DECLARE_NOT_AGGREGATABLE(CCollFac)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CCollFac)
	COM_INTERFACE_ENTRY(ICollFac)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)	
	COM_INTERFACE_ENTRY(IPersistStorage)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)
	COM_INTERFACE_ENTRY(IClonable)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ICollFac
public:
    struct TFndItDta_TAMP
	 {
       TFndItDta_TAMP( IFactor* p )
		{
		  m_p = p;
		}

	   bool operator()( const TMAP::value_type& rI )
		{
          return rI.second == m_p;
		}

	   IFactor* m_p;
	 };

    STDMETHOD(GetIndexOfItem)(/*[in]*/IFactor* pEN, /*[out, retval]*/BSTR* pkey)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


       ATLTRACE2(atlTraceCOM, 0, _T("ICollLingvo::GetIndexOfItem\n"));	   

       if( pEN == NULL || pkey == NULL ) return E_POINTER;

	   TFndItDta_TAMP dta( pEN );
	   TMAP::iterator it = find_if( m_coll.begin(), m_coll.end(), dta );
	   if( it != m_coll.end() )
		{
          *pkey = it->first.copy();
		  return S_OK;
		}
       
	   std::basic_stringstream<WCHAR> strm;
	   strm << std::setbase(16);
	   strm << L"ICollLingvo::GetIndexOfItem: " << pEN << L" - item not found";
	   Error( strm.str().c_str(), IID_ICollFac, E_FAIL );   
	   return E_FAIL;
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
	   TMAP::iterator itRes = find_if( m_coll.begin(), m_coll.end(), IsDrt_CCollFac );

	   return itRes != m_coll.end() ? S_OK:S_FALSE;
	 }
	
	STDMETHOD(InitNew)(IStorage*)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::InitNew\n"));
		RemoveAll();		

		m_bRequiresSave = true;

		return S_OK;
	}
	STDMETHOD(Load)(IStorage* pStorage)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	  ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::Load\n"));

	  RemoveAll();

	  CComPtr<IStream> spInfoStream;
	  HRESULT hr = pStorage->OpenStream( L"Content", NULL, 
	    STGM_READ|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, &spInfoStream );
	  if( FAILED(hr) ) 
       {         
		 PassError( L"IStorage.OpenStream", hr, GetObjectCLSID(), IID_ICollFac  );   
		 return hr;
	   }
	  CComVariant cv;
	  hr = cv.ReadFromStream( spInfoStream );
	  if( FAILED(hr) )
	   {
	     PassError( L"ReadFromStream", hr, GetObjectCLSID(), IID_ICollFac  );   
	     return hr;
	   }
	  if( cv.vt != VT_I4 )	   
	   {
	     PassError( L"Bad format", E_FAIL, GetObjectCLSID(), IID_ICollFac  );   
	     return E_FAIL;
	   }

	  for( long i = 0; i < cv.lVal; ++i )
	   {
         CComBSTR bsKey;
		 hr = bsKey.ReadFromStream( spInfoStream );
		 if( FAILED(hr) )
		  {
            PassError( L"ReadFromStream", hr, GetObjectCLSID(), IID_ICollFac  );   			   
	        return hr;
		  }

         CComObject<CFactor>* pCF;
		 hr = CComObject<CFactor>::CreateInstance( &pCF );
		 if( FAILED(hr) )
		  {
            PassError( L"CComObject<CFactor>::CreateInstance", hr, GetObjectCLSID(), IID_ICollFac  );   			   
	        return hr;
		  }

		 CComPtr<IPersistStreamInit> spPSI;
		 hr = pCF->_InternalQueryInterface( IID_IPersistStreamInit, (void**)&spPSI );
		 if( FAILED(hr) )
		  {
            PassError( L"_InternalQueryInterface", hr, GetObjectCLSID(), IID_ICollFac  );   			   
			delete pCF;
	        return hr;
		  }

		 hr = spPSI->Load( spInfoStream );
		 if( FAILED(hr) ) 
		  {			   
	        PassError( L"IStreamInitNew.Load", hr, GetObjectCLSID(), IID_ICollFac  );   			   
			spPSI.Release(); 
			//delete pLE;
	        return hr;
		  }

		 CComQIPtr<IFactor> spLEnum( spPSI );
		 spPSI.Release(); 

		 m_coll.insert( TMAP::value_type(Chk(bsKey), spLEnum) );
		 
		 spLEnum.Detach();
	   }

	  m_bRequiresSave = false;

	  return S_OK;
	}    
	

  STDMETHOD(Save)(IStorage* pStorage, BOOL fSameAsLoad)
	{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	  ATLTRACE2(atlTraceCOM, 0, _T("IPersistStorage::Save\n"));

	  if( !pStorage ) return E_POINTER;
      
	  CComPtr<IStream> spInfoStream;
	  HRESULT hr = pStorage->CreateStream( L"Content", 
	    STGM_CREATE|STGM_READWRITE|STGM_DIRECT|STGM_SHARE_EXCLUSIVE, 0, 0, &spInfoStream );
	  if( FAILED(hr) ) 
       {         
		 PassError( L"IStorage.OpenStream", hr, GetObjectCLSID(), IID_ICollFac  );   
		 return hr;
	   }
	  CComVariant cv( (long)m_coll.size() );
	  hr = cv.WriteToStream( spInfoStream );
	  if( FAILED(hr) )
	   {
	     PassError( L"WriteToStream", hr, GetObjectCLSID(), IID_ICollFac  );   
	     return hr;
	   }
	  

	  TMAP::iterator itSta( m_coll.begin() );
	  TMAP::iterator itEnd( m_coll.end() );
	  for( ; itSta != itEnd; ++itSta )
	   {
         CComBSTR bsKey( (BSTR)itSta->first );
		 
		 hr = bsKey.WriteToStream( spInfoStream );
		 if( FAILED(hr) )
		  {
            PassError( L"WriteToStream", hr, GetObjectCLSID(), IID_ICollFac  );   			   
	        return hr;
		  }

         CComQIPtr<IPersistStreamInit> spPSI( itSta->second );
		 if( !spPSI )
		  {
		    PassError( L"CComQIPtr<IPersistStreamInit>", E_FAIL, GetObjectCLSID(), IID_ICollFac  );   			   
			return E_FAIL;
		  }

		 hr = spPSI->Save( spInfoStream, TRUE );
		 if( FAILED(hr) ) 
		  {			   
	        PassError( L"IStreamInitNew.Save", hr, GetObjectCLSID(), IID_ICollFac  );   			   
			spPSI.Release(); 			
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

public:
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   HRESULT hr = TCloneImpl2::Clone( ppUnk );

	   if( SUCCEEDED(hr) )	    
		m_pObjTmp->m_bRequiresSave = false;
			
		
	   return hr;
	 }

public:
  CComObject<CCollFac>* m_pObjTmp; //for CIClonableImpl_CollOnMap

};

#endif //__COLLFAC_H_
