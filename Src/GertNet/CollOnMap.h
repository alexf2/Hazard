// CollLingvo.h : Declaration of the CCollLingvo

#ifndef __COLLLONMAP_H_
#define __COLLLONMAP_H_

#include "resource.h"       // main symbols
#include <map>
#include <comutil.h>
#include <comdef.h>

#include "PassErr.h"
//#include <OtlComCollectionSTL.h>

//TYPE:    interface (ILingvoEnum)
//ITEMMAP: type of key (long)
//BASE:    baseclass

template<class TYPE, class ITEMMAP, class BASE>
class CComCollectionOnMap: public BASE   
{

public:

    typedef std::map<ITEMMAP, TYPE*, std::less<ITEMMAP> > TMAP;
	typedef ITEMMAP ITEMMAP_B;

	CComCollectionOnMap():m_dwMaxKey(0), m_bRequiresSave(false) {}
    virtual ~CComCollectionOnMap() { RemoveAll(); }

	// standard Collection methods.
	STDMETHOD(Add)(TYPE *pIType, ITEMMAP index, ITEMMAP* indexAss);
	STDMETHOD(get_Count)(long *pVal);
	STDMETHOD(Item)(ITEMMAP index, TYPE **pvar);
	STDMETHOD(get__NewEnum)(LPUNKNOWN *pVal);
	STDMETHOD(Remove)(ITEMMAP index);
	STDMETHOD(Clear)();

    typedef CComEnum<CComIEnum<VARIANT>, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT>, CComSingleThreadModel> _CComCollEnum;

protected:
    /*long AddItem(TYPE *pIType, BSTR bstrKey);
	HRESULT RemoveItem(long lIndex);
	HRESULT RemoveItem(BSTR bstrKey);
    TYPE* Item(long lIndex);
    TYPE* Item(BSTR bstrKey);*/

public:
    TMAP m_coll;
protected:
    
	DWORD m_dwMaxKey;

	bool m_bRequiresSave;

	bool RemoveAll();
 };

template<class TYPE, class ITEMMAP, class BASE>
STDMETHODIMP CComCollectionOnMap<TYPE,ITEMMAP,BASE>::Clear()
 {
   RemoveAll();
   return S_OK;
 }

template<class TYPE, class ITEMMAP, class BASE>
STDMETHODIMP CComCollectionOnMap<TYPE,ITEMMAP,BASE>::Add( TYPE *pIType, ITEMMAP index, ITEMMAP* indexAss )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    HRESULT hr = S_OK;

    if( pIType != NULL )
     {
		// Validate the interface TYPE support.
		TYPE *pI=NULL;
		hr = pIType->QueryInterface( _uuidof(TYPE), (void **)&pI );
		pI->Release();

		if(SUCCEEDED(hr))
		 {		   
		   bool bFlAss = (index == -1);
		   if( bFlAss ) index = ++m_dwMaxKey;
		   else m_dwMaxKey = max( (ITEMMAP)m_dwMaxKey, index );
		   if( m_coll.insert(TMAP::value_type(index, pIType)).second == false )
			{
			  PassError( L"Duplicated index", E_FAIL, GUID_NULL, GUID_NULL  );   
			  return E_FAIL;
			}
		   if( indexAss != NULL ) *indexAss = index;
		   if( bFlAss )
			{
			  CComQIPtr<ILongKey> spLK( pIType );
			  if( spLK.p != NULL ) spLK->Set( index );
			}
		   m_bRequiresSave = true;

		   pIType->AddRef();
		 }
		else
		 {
			hr = E_NOINTERFACE;
		 }
     }
    else
     {
        hr = E_POINTER;
     }
    return hr;
 }


template<class TYPE, class ITEMMAP, class BASE>
STDMETHODIMP CComCollectionOnMap<TYPE,ITEMMAP,BASE>::get_Count( long *pVal )
 { 
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    HRESULT hr = S_OK;
    if( pVal != NULL )
		*pVal = m_coll.size();
    else
        hr = E_POINTER;

	return hr; 
 }

template<class TYPE, class ITEMMAP, class BASE>
STDMETHODIMP CComCollectionOnMap<TYPE,ITEMMAP,BASE>::Item( ITEMMAP index, TYPE **pvar )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    HRESULT hr = S_OK;
    if( pvar != NULL )
     {
	    TMAP::iterator it = m_coll.find( index );
		if( it != m_coll.end() )
          *pvar = it->second, it->second->AddRef();
		else
		 {
		   hr = E_INVALIDARG;         
		   std::basic_stringstream<WCHAR> strm;
		   strm << L"get_Item:\nBad key value: " << index;
		   PassError( strm.str().c_str(), E_INVALIDARG, GUID_NULL, GUID_NULL  );
		 }
     }
    else
     {
        hr = E_POINTER;
     }

	return hr; 
 }


template<class TYPE, class ITEMMAP, class BASE>
STDMETHODIMP CComCollectionOnMap<TYPE,ITEMMAP,BASE>::get__NewEnum( LPUNKNOWN *ppVal )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    HRESULT hr = S_OK;

    if( ppVal == NULL )
	    return E_POINTER;
    *ppVal = NULL;

    _CComCollEnum *pEnum = NULL;
    ATLTRY( pEnum = new CComObject<_CComCollEnum> )
    if( pEnum == NULL )
	    return E_OUTOFMEMORY;

    // allocate an initialize a vector of VARIANT objects
	int nCount = m_coll.size();
	unsigned int uiSize = sizeof(VARIANT) * nCount;
    VARIANT* pVar = (VARIANT*)new VARIANT[ nCount ];
	memset( pVar, '\0', uiSize );

	// Set the VARIANTs that will hold the TYPE Interface pointers.
    int i = 0;
	TMAP::iterator it;
	for( it = m_coll.begin(); it != m_coll.end(); it++ )
     {
		TYPE* pIP = NULL;
		pVar[i].vt = VT_DISPATCH;
		pIP = (*it).second;
		if( pIP != NULL )
		 {
			hr = pIP->QueryInterface( IID_IDispatch, (void**)&pVar[i].pdispVal );
			pIP = NULL;
		 }
		i++;
     }

    // copy the array of VARIANTs
	hr = pEnum->Init( &pVar[0],
					  &pVar[nCount],
					  reinterpret_cast<TYPE*>(this)/*NULL*/, AtlFlagTakeOwnership );
    if( SUCCEEDED(hr) )     	 
	  hr = pEnum->QueryInterface( IID_IUnknown, (void**)ppVal );
			 

	if( FAILED(hr) )
	 {
	   delete pEnum;	 
	   // Iterate through the VARIANT pointer array clearing each VARIANT
	   i = 0;
	   while ( i < nCount )
		{
		   VariantClear( &pVar[i] );
		   i++;
		}
	   delete pVar;
	 }
	

	return hr; 
 }

template<class TYPE, class ITEMMAP, class BASE>
STDMETHODIMP CComCollectionOnMap<TYPE,ITEMMAP,BASE>::Remove( ITEMMAP index )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    //HRESULT hr;
    
	TMAP::iterator it = m_coll.find( index );
	if( it != m_coll.end() )
	 {
	    TYPE* pType = (*it).second;
		//ATLTRACE("Remove: %S:%0x\n", bstrKey, pType);
		if( pType )
		  pType->Release();
		m_coll.erase( it );

		m_bRequiresSave = true;

		return S_OK;
	 }	
	else
	 {
	   std::basic_stringstream<WCHAR> strm;
	   strm << L"Remove:\nBad key value: " << index;
	   PassError( strm.str().c_str(), E_INVALIDARG, GUID_NULL, GUID_NULL  );

	   return E_INVALIDARG;
	 }	
 }

template<class TYPE, class ITEMMAP, class BASE>
bool CComCollectionOnMap<TYPE,ITEMMAP,BASE>::RemoveAll()
 {   
	TMAP::iterator it;
	TYPE* pType = NULL;
	for( it = m_coll.begin(); it != m_coll.end(); it++ )
	 {
		pType = (*it).second;
		if( pType )
		  pType->Release();
	 }
	m_coll.erase( m_coll.begin(), m_coll.end() );

    m_bRequiresSave = true;

	return true;
 }

//----------------------------------------

template<class TYPE, class BASE>
class CComCollectionOnMap_BSTR: public BASE   
{

public:

    typedef std::map<_bstr_t, TYPE*, std::less<_bstr_t> > TMAP;

	CComCollectionOnMap_BSTR():m_bRequiresSave(false) {}
    virtual ~CComCollectionOnMap_BSTR() { RemoveAll(); }

	// standard Collection methods.
	STDMETHOD(Add)(TYPE *pIType, BSTR key);
	STDMETHOD(get_Count)(long *pVal);
	STDMETHOD(Item)(BSTR key, TYPE **pvar);
	STDMETHOD(get__NewEnum)(LPUNKNOWN *pVal);
	STDMETHOD(Remove)(BSTR key);
	STDMETHOD(Clear)();

    typedef CComEnum<CComIEnum<VARIANT>, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT>, CComSingleThreadModel> _CComCollEnum;
    
public:
    TMAP m_coll;	
protected:
    

	bool m_bRequiresSave;

	bool RemoveAll();
 };

template<class TYPE, class BASE>
STDMETHODIMP CComCollectionOnMap_BSTR<TYPE,BASE>::Clear()
 {
   RemoveAll();
   return S_OK;
 }


template<class TYPE, class BASE>
STDMETHODIMP CComCollectionOnMap_BSTR<TYPE,BASE>::Add( TYPE *pIType, BSTR key )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    HRESULT hr = S_OK;

    if( pIType != NULL )
     {
		// Validate the interface TYPE support.
		TYPE *pI=NULL;
		hr = pIType->QueryInterface( _uuidof(TYPE), (void **)&pI );
		pI->Release();

		if(SUCCEEDED(hr))
		 {		   		   
		   if( m_coll.insert(TMAP::value_type(key, pIType)).second == false )
			{
			  PassError( L"Duplicated index", E_FAIL, GUID_NULL, GUID_NULL  );   
			  return E_FAIL;
			}
		   
		   CComQIPtr<IBSTRKey> spBSK( pIType );
		   if( spBSK.p != NULL ) spBSK->Set( key );
		   m_bRequiresSave = true;

		   pIType->AddRef();
		 }
		else
		 {
			hr = E_NOINTERFACE;
		 }
     }
    else
     {
        hr = E_POINTER;
     }
    return hr;
 }


template<class TYPE, class BASE>
STDMETHODIMP CComCollectionOnMap_BSTR<TYPE,BASE>::get_Count( long *pVal )
 { 
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    HRESULT hr = S_OK;
    if( pVal != NULL )
		*pVal = m_coll.size();
    else
        hr = E_POINTER;

	return hr; 
 }

template<class TYPE, class BASE>
STDMETHODIMP CComCollectionOnMap_BSTR<TYPE,BASE>::Item( BSTR key, TYPE **pvar )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	_bstr_t bsKey( key, true );

    HRESULT hr = S_OK;
    if( pvar != NULL && key != NULL )
     {
	    TMAP::iterator it = m_coll.find( bsKey );
		if( it != m_coll.end() )
          *pvar = it->second, it->second->AddRef();
		else
		 {
		   hr = E_INVALIDARG;         
		   std::basic_stringstream<WCHAR> strm;
		   strm << L"get_Item:\nBad key value: " << (wchar_t*)bsKey;
		   PassError( strm.str().c_str(), E_INVALIDARG, GUID_NULL, GUID_NULL  );
		 }
     }
    else
     {
        hr = E_POINTER;
     }

	return hr; 
 }

template<class TYPE,  class BASE>
STDMETHODIMP CComCollectionOnMap_BSTR<TYPE,BASE>::get__NewEnum( LPUNKNOWN *ppVal )
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
	   hr = pEnum->Init( &pVar[0],
						 &pVar[nCount],
						 static_cast<IUnknown*>(this)/*NULL*/, AtlFlagTakeOwnership );
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

template<class TYPE,  class BASE>
STDMETHODIMP CComCollectionOnMap_BSTR<TYPE,BASE>::Remove( BSTR bsKey )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    //HRESULT hr;
    
	TMAP::iterator it = m_coll.find( bsKey );
	if( it != m_coll.end() )
	 {
	    TYPE* pType = (*it).second;
		//ATLTRACE("Remove: %S:%0x\n", bstrKey, pType);
		if( pType )
		  pType->Release();
		m_coll.erase( it );
		m_bRequiresSave = true;

		return S_OK;
	 }	
	else
	 {
	   std::basic_stringstream<WCHAR> strm;
	   strm << L"Remove:\nBad key value: " << (BSTR)Chk(bsKey);
	   PassError( strm.str().c_str(), E_INVALIDARG, GUID_NULL, GUID_NULL  );

	   return E_INVALIDARG;
	 }	
 }

template<class TYPE, class BASE>
bool CComCollectionOnMap_BSTR<TYPE,BASE>::RemoveAll()
 {   
	TMAP::iterator it;
	TYPE* pType = NULL;
	for( it = m_coll.begin(); it != m_coll.end(); it++ )
	 {
		pType = (*it).second;
		if( pType )
		  pType->Release();
	 }
	m_coll.erase( m_coll.begin(), m_coll.end() );
	m_bRequiresSave = true;

	return true;
 }

//
template<class IMPLCLASS, class MAP_TYP, class INTF_COLLITEM>
class CIClonableImpl_CollOnMap: public IClonable 
 {
public:

    CIClonableImpl_CollOnMap() {}

// IClonable
    STDMETHOD(Clone)(/*[out, retval]*/ IUnknown** ppUnk)
	 {	   
	   if( ppUnk == NULL ) return E_POINTER;

        IMPLCLASS* pIC = static_cast<IMPLCLASS*>(this);

		CComObject<IMPLCLASS>* pObj;
		HRESULT hr = CComObject<IMPLCLASS>::CreateInstance( &pObj );
		if( FAILED(hr) )
		 {
		   PassError( L"IClonable::Clone: CComObject<>::CreateInstance", hr, pIC->GetObjectCLSID(), IID_IClonable );
		   return hr;
		 }

		pIC->m_pObjTmp = pObj;

		CComPtr<IUnknown> spTmp;
		hr = pObj->_InternalQueryInterface( IID_IUnknown, (void**)&spTmp );
		if( FAILED(hr) )
		 {
		   pIC->m_pObjTmp = NULL;
		   delete pObj;
		   PassError( L"IClonable::Clone: CComObject<>::_InternalQueryInterface", hr, pIC->GetObjectCLSID(), IID_IClonable );
		   return hr;
		 }

		MAP_TYP::iterator it1( pIC->m_coll.begin() );
		MAP_TYP::iterator it2( pIC->m_coll.end() );
		for( ; it1 != it2; ++it1 )
		 {
		   CComQIPtr<IClonable> spClon( it1->second );
		   if( !spClon )
			{
			  PassError( L"IClonable::Clone: CComQIPtr<IClonable>", E_FAIL, pIC->GetObjectCLSID(), IID_IClonable );
		      return E_FAIL;
			}
		   CComPtr<IUnknown> spCopy;
		   hr = spClon->Clone( &spCopy );
		   if( FAILED(hr) )
			{			 
			  PassError( L"IClonable::Clone: spClon->Clone", hr, pIC->GetObjectCLSID(), IID_IClonable );
			  return hr;
			}
		   
		   CComQIPtr<INTF_COLLITEM> spFac( spCopy.p );
		   spFac.p->AddRef();
		   pObj->m_coll.insert( MAP_TYP::value_type(it1->first, spFac.p) );
		 }
	    				

		spTmp.CopyTo( ppUnk );	

		return S_OK;
	 }

 };

#endif
