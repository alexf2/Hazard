// ICollFChange.h : Declaration of the CICollFChange

#ifndef __ICOLLFCHANGE_H_
#define __ICOLLFCHANGE_H_

#include "resource.h"       // main symbols

#include "CollOnMap.h"
#include "PassErr.h"
#include "FChange.h"

#include <algorithm>
#include <list>
#include <sstream>
#include <iomanip>
using namespace std;


typedef  CComCollectionOnMap< IFChange, 
	  long,
	  IDispatchImpl<IICollFChange, &IID_IICollFChange, &LIBID_GERTNETLib> > CCL3;

inline int IsDrt_CCollIFC( CCL3::TMAP::value_type& rVal )
 {
   if( rVal.second == NULL ) return false;
   CComQIPtr<IPersistStreamInit> spSI( rVal.second );
   if( !spSI ) return false;	
   else return spSI->IsDirty() == S_OK ? true:false;
 }

class CICollFChange;
typedef CIClonableImpl_CollOnMap<CICollFChange, CCL3::TMAP, IFChange> TCloneImpl3;

/////////////////////////////////////////////////////////////////////////////
// CICollFChange
class ATL_NO_VTABLE CICollFChange : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CICollFChange, &CLSID_ICollFChange>,	
	public CCL3,
	public IPersistStreamInit,
	public ISupportErrorInfo,
	public TCloneImpl3
	//public IClonable
{
public:
	CICollFChange()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_ICOLLFCHANGE)
DECLARE_NOT_AGGREGATABLE(CICollFChange)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CICollFChange)
	COM_INTERFACE_ENTRY(IICollFChange)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IClonable)
	COM_INTERFACE_ENTRY(IPersistStreamInit)
	COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IICollFChange
public:

 // ICollLingvo
public:
	STDMETHOD(LookNextIndex)(/*[out, retval]*/long* plIndex)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		
	   ATLTRACE2(atlTraceCOM, 0, _T("IICollFChange::LookNextIndex\n"));	   

       if( plIndex == NULL ) return E_POINTER;
	   *plIndex = m_dwMaxKey + 1;
	   return S_OK;
	 }

	STDMETHOD(GetClassID)(CLSID *pClassID)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::GetClassID\n"));
		*pClassID = GetObjectCLSID();
		return S_OK;
	 }	
	STDMETHOD(IsDirty)()
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::IsDirty\n"));		
	   
	   if( m_bRequiresSave ) return S_OK;
   
	   bool bFlFnd = false;   	   
	   TMAP::iterator itRes = find_if( m_coll.begin(), m_coll.end(), IsDrt_CCollIFC );

	   return itRes != m_coll.end() ? S_OK:S_FALSE;
	 }

	STDMETHOD(Load)(LPSTREAM pStm)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Load\n"));

	   if( !pStm ) return E_POINTER;
      	   
	   CComVariant cv;
	   HRESULT hr = cv.ReadFromStream( pStm );
	   if( FAILED(hr) )
		{
		  PassError( L"ReadFromStream", hr, GetObjectCLSID(), IID_IICollFChange  );   
		  return hr;
		}
	   if( cv.vt != VT_I4 )
		{
		  PassError( L"Illegal stream format", E_FAIL, GetObjectCLSID(), IID_IICollFChange  );   
		  return E_FAIL;
		}
	   hr = pStm->Read( &m_dwMaxKey, sizeof(m_dwMaxKey), NULL );
	   if( FAILED(hr) )
		{
		  PassError( L"Read", hr, GetObjectCLSID(), IID_IICollFChange  );   
		  return hr;
		}
	   
	   	   
	   for( int i = cv.lVal; i > 0; --i )
		{
		  CComObject<CFChange>* pLE;
		  hr = CComObject<CFChange>::CreateInstance( &pLE );
		  if( FAILED(hr) )
		   {
             PassError( L"CComObject<CFChange>::CreateInstance", hr, GetObjectCLSID(), IID_IICollFChange  );   			   
	         return hr;
		   }

		  CComPtr<IPersistStreamInit> spPSI;
		  hr = pLE->_InternalQueryInterface( IID_IPersistStreamInit, (void**)&spPSI );
		  if( FAILED(hr) )
		   {
             PassError( L"_InternalQueryInterface", hr, GetObjectCLSID(), IID_IICollFChange  );   			   
			 delete pLE;
	         return hr;
		   }		  

		  CCL3::ITEMMAP_B imKey;
		  hr = pStm->Read( &imKey, sizeof(imKey), NULL );
		  if( FAILED(hr) )
		   {
			 PassError( L"Read", hr, GetObjectCLSID(), IID_IICollFChange  );   
			 return hr;
		   }

		  hr = spPSI->Load( pStm );
		  if( FAILED(hr) ) 
		   {			   
	         PassError( L"IStreamInitNew.Load", hr, GetObjectCLSID(), IID_IICollFChange  );   			   
			 spPSI.Release(); 
			 //delete pLE;
	         return hr;
		   }

		  CComQIPtr<IFChange> spLEnum( spPSI );
		  spPSI.Release(); 
		  /*CComQIPtr<ILongKey> spKey( spPSI );
		  
		  long lK;
		  spKey->Get( &lK );*/

		  m_coll.insert( TMAP::value_type(imKey, spLEnum) );		  
		  spLEnum.Detach();
		}

	   m_bRequiresSave = false;

	   return S_OK;
	 }

	STDMETHOD(Save)(LPSTREAM pStm, BOOL fClearDirty)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Save\n"));

	   if( !pStm ) return E_POINTER;
      	   
	   CComVariant cv( (long)m_coll.size() );
	   HRESULT hr = cv.WriteToStream( pStm );
	   if( FAILED(hr) )
		{
		  PassError( L"WriteToStream", hr, GetObjectCLSID(), IID_IICollFChange  );   
		  return hr;
		}
	   hr = pStm->Write( &m_dwMaxKey, sizeof(m_dwMaxKey), NULL );
	   if( FAILED(hr) )
		{
		  PassError( L"WriteToStream", hr, GetObjectCLSID(), IID_IICollFChange  );   
		  return hr;
		}
	   
	   
	   TMAP::iterator itSta( m_coll.begin() );
	   TMAP::iterator itEnd( m_coll.end() );
	   for( ; itSta != itEnd; ++itSta )
		{
		  hr = pStm->Write( &itSta->first, sizeof(itSta->first), NULL );
		  if( FAILED(hr) )
		   {
			 PassError( L"WriteToStream", hr, GetObjectCLSID(), IID_IICollFChange  );   			   
			 return hr;
		   }

		  CComQIPtr<IPersistStreamInit> spPSI( itSta->second );
		  if( !spPSI )
		   {
			 PassError( L"CComQIPtr<IPersistStreamInit>", E_FAIL, GetObjectCLSID(), IID_IICollFChange  );   			   
			 return E_FAIL;
		   }

		  hr = spPSI->Save( pStm, TRUE );
		  if( FAILED(hr) ) 
		   {			   
			 PassError( L"IStreamInitNew.Save", hr, GetObjectCLSID(), IID_IICollFChange  );   			   
			 spPSI.Release(); 			
			 return hr;
		   }		 
		}

	   if( fClearDirty == TRUE )
	     m_bRequiresSave = false;

	   return S_OK;
	 }
	
	STDMETHOD(GetSizeMax)(ULARGE_INTEGER FAR* /* pcbSize */)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

		ATLTRACENOTIMPL(_T("IPersistStreamInit::GetSizeMax"));
		return E_NOTIMPL;
	 }
	
	STDMETHOD(InitNew)()
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::InitNew\n"));
	   RemoveAll();		

	   m_bRequiresSave = true;

	   return S_OK;
	 }

public:
// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk)
	 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	   HRESULT hr = TCloneImpl3::Clone( ppUnk );

	   if( SUCCEEDED(hr) )	    	    
		m_pObjTmp->m_dwMaxKey = m_dwMaxKey;
		
       return hr;		
	 }

public:
  CComObject<CICollFChange>* m_pObjTmp; //for CIClonableImpl_CollOnMap

};

#endif //__ICOLLFCHANGE_H_
