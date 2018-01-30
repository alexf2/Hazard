// CollSF.cpp : Implementation of CCollSF
#include "stdafx.h"
#include "GertNet.h"
#include "CollSF.h"

/////////////////////////////////////////////////////////////////////////////
// CCollSF

#include <set>
#include <algorithm>
using namespace std;


inline int TestZeroDelta( CCollSF::TMAP::value_type& rVal )
 {
   CComObject<CSafetyPrecaution>* pSaf = 
	  static_cast< CComObject<CSafetyPrecaution>* >( rVal.second );
   return pSaf->DeltaRef() == 0;
 }

inline void ZeroDelta( CCollSF::TMAP::value_type& rVal )
 {
   CComObject<CSafetyPrecaution>* pSaf = 
	  static_cast< CComObject<CSafetyPrecaution>* >( rVal.second );
   pSaf->DeltaRef() = 0;
 }

STDMETHODIMP CCollSF::get_IsValidDelta(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	TMAP::iterator it = find_if( m_coll.begin(), m_coll.end(), TestZeroDelta );
	if( it == m_coll.end() ) *pVal = VARIANT_TRUE;
	else *pVal = VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CCollSF::InvalidateDelta()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	for_each( m_coll.begin(), m_coll.end(), ZeroDelta );

	return S_OK;
}

STDMETHODIMP CCollSF::get_Sorting(CollSFSorting *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_csfSorting;

	return S_OK;
}

STDMETHODIMP CCollSF::put_Sorting(CollSFSorting newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_csfSorting = newVal;

	return S_OK;
}

STDMETHODIMP CCollSF::get_IsDirty(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = (IsDirty() == S_OK ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CCollSF::get_Profit(CURRENCY *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_ocProfit;

	return S_OK;
}

STDMETHODIMP CCollSF::put_Profit(CURRENCY newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	try {
	   m_ocProfit = newVal;
	 }
	catch( CException* pE )
	 {
	   pE->Delete();
       PassError( L"CSafetyPrecaution::put_Profit: ", E_INVALIDARG, GetObjectCLSID(), IID_ISafetyPrecaution );
	   return E_INVALIDARG;
	 }
	

	return S_OK;
}

STDMETHODIMP CCollSF::get_Ke(double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_dKe;

	return S_OK;
}

STDMETHODIMP CCollSF::put_Ke(double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_dKe = newVal;

	return S_OK;
}

STDMETHODIMP CCollSF::get_Cost(CURRENCY *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_ocCost;

	return S_OK;
}

STDMETHODIMP CCollSF::put_Cost(CURRENCY newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	try {
	   m_ocCost = newVal;
	 }
	catch( CException* pE )
	 {
	   pE->Delete();
       PassError( L"CSafetyPrecaution::put_Cost: ", E_INVALIDARG, GetObjectCLSID(), IID_ISafetyPrecaution );
	   return E_INVALIDARG;
	 }
	

	return S_OK;
}

STDMETHODIMP CCollSF::get_DeltaQ(double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_dDeltaQ;

	return S_OK;
}

STDMETHODIMP CCollSF::put_DeltaQ(double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_dDeltaQ = newVal;

	return S_OK;
}

STDMETHODIMP CCollSF::CheckCompatible(ICollSF** pSFNone, VARIANT_BOOL* bResult)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( !bResult ) return E_POINTER;

   HRESULT hr = S_OK;

   set<long> setSelectedSF, setNonCompatibleSF, setInt;
   TMAP::iterator it( m_coll.begin() );
   for( ; it != m_coll.end(); ++it )
	{
	  CComObject<CSafetyPrecaution>& rSF = *static_cast<CComObject<CSafetyPrecaution>*>( it->second );
	  setSelectedSF.insert( rSF.m_lID );
	  setNonCompatibleSF.insert( rSF.m_vecCollNC.begin(), rSF.m_vecCollNC.end() );
	}

   set_intersection( setSelectedSF.begin(), setSelectedSF.end(), 
	  setNonCompatibleSF.begin(), setNonCompatibleSF.end(), inserter(setInt, setInt.begin()) );
   
   *bResult = (setInt.size() != 0 ? VARIANT_FALSE:VARIANT_TRUE);

   if( pSFNone )
	{
      CComObject<CCollSF>* pObj;
	  hr = CComObject<CCollSF>::CreateInstance( &pObj );
	  if( FAILED(hr) )
	   {
		 PassError( L"CCollSF::CheckCompatible: CComObject<CCollSF>::CreateInstance", hr, GetObjectCLSID(), IID_ICollSF );
		 return hr;
	   }

	  CComPtr<ICollSF> spTmp;
	  hr = pObj->_InternalQueryInterface( IID_ICollSF, (void**)&spTmp );
	  if( FAILED(hr) )
	   {
		 delete pObj;
		 PassError( L"CCollSF::CheckCompatible: CComObject<CCollSF>::_InternalQueryInterface", hr, GetObjectCLSID(), IID_ICollSF );
		 return hr;
	   }

	  CComQIPtr<IPersistStorage> spStor( spTmp );
	  spStor->InitNew( NULL );

	  hr = spTmp.CopyTo( pSFNone );
	  if( SUCCEEDED(hr) )
	   {
	     set<long>::iterator itSta( setInt.begin() );
		 for( ; itSta != setInt.end(); ++itSta )
		  {
		    ISafetyPrecaution *pSf = m_coll.find(*itSta)->second;
			ATLASSERT( pSf != NULL );
		    ATLASSERT( pObj->m_coll.insert(TMAP::value_type( *itSta, pSf )).second == true );
			pSf->AddRef();
		  }
		 pObj->m_dwMaxKey = m_dwMaxKey;
	   }
	}

   return hr;
}
