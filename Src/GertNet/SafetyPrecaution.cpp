// SafetyPrecaution.cpp : Implementation of CSafetyPrecaution
#include "stdafx.h"
#include "GertNet.h"
#include "SafetyPrecaution.h"
#include "ICollFChange.h"

#include "PassErr.h"

#include <tchar.h>


#include <sstream>
#include <string>
#include <iomanip>
using namespace std;


/////////////////////////////////////////////////////////////////////////////
// CSafetyPrecaution

STDMETHODIMP CSafetyPrecaution::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ISafetyPrecaution,
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

STDMETHODIMP CSafetyPrecaution::get_Name(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	return m_bsName.CopyTo( pVal );	
}


STDMETHODIMP CSafetyPrecaution::put_Name(BSTR newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_bsName = newVal;
	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::get_Cost(CURRENCY *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_ocCost;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::put_Cost(CURRENCY newVal)
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

	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::get_DeltaQ(double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_dDeltaQ;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::put_DeltaQ(double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_dDeltaQ = newVal;
	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::Clone(/*[out, retval]*/ IUnknown** ppUnk)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( ppUnk == NULL ) return E_POINTER;

    CComObject<CSafetyPrecaution>* pObj;
	HRESULT hr = CComObject<CSafetyPrecaution>::CreateInstance( &pObj );
	if( FAILED(hr) )
	 {
       PassError( L"CSafetyPrecaution::Copy: CComObject<CSafetyPrecaution>::CreateInstance", hr, GetObjectCLSID(), IID_ISafetyPrecaution );
       return hr;
	 }

	CComPtr<IUnknown> spTmp;
	hr = pObj->_InternalQueryInterface( IID_IUnknown, (void**)&spTmp );
	if( FAILED(hr) )
	 {
	   delete pObj;
       PassError( L"CSafetyPrecaution::Copy: CComObject<CSafetyPrecaution>::_InternalQueryInterface", hr, GetObjectCLSID(), IID_ISafetyPrecaution );
       return hr;
	 }

	pObj->m_lID = m_lID;
	pObj->m_dDeltaQ = m_dDeltaQ;
	pObj->m_ocCost = m_ocCost;
	pObj->m_bsName = m_bsName;

	pObj->m_ocProfit = m_ocProfit;
	pObj->m_dKe = m_dKe;

	
	CComQIPtr<IClonable> spClon( m_spCollFChange.p );
	if( !spClon.p )
	 {	   
	   Error( L"CSafetyPrecaution::Copy: CComQIPtr<IClonable>", IID_ISafetyPrecaution, E_FAIL );   
       return E_FAIL;
	 }
	CComPtr<IUnknown> spUnkCopy;
	hr = spClon->Clone( &spUnkCopy );
	if( FAILED(hr) )
	 {
       PassError( L"CSafetyPrecaution::Copy: cann't clone", hr, GetObjectCLSID(), IID_ISafetyPrecaution );
       return hr;
	 }
	pObj->m_spCollFChange.Release();
	hr = spUnkCopy.QueryInterface( &pObj->m_spCollFChange.p );
	if( FAILED(hr) )
	 {
       PassError( L"CSafetyPrecaution::Copy: QueryInterface", hr, GetObjectCLSID(), IID_ISafetyPrecaution );
       return hr;
	 }
		
	
    pObj->m_vecCollNC = m_vecCollNC;
	pObj->m_bRequiresSave = false;

	
	spTmp.CopyTo( ppUnk );

	return S_OK;
 }

STDMETHODIMP CSafetyPrecaution::get_ID(VARIANT *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;

	VariantInit( pVal );
	pVal->vt = VT_I4;
	pVal->lVal = m_lID;


	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::get_FChanges(IICollFChange **pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	
	return m_spCollFChange.CopyTo( pVal );
}

HRESULT CSafetyPrecaution::InternalIN()
 {   
   m_lID = 0;

   m_dDeltaQ = 0;
   CY cy; cy.int64 = 0;
   m_ocCost = cy;
   m_bsName = L"<Not assigned>";

   m_ocProfit = cy;
   m_dKe = 0;
     
   HRESULT hr = InitCollections();     
   if( SUCCEEDED(hr) )
	{
	  CComQIPtr<IPersistStreamInit> sp( m_spCollFChange.p );
	  if( sp.p )
	    hr = sp->InitNew();	  
	  else hr = E_NOINTERFACE;

	  m_vecCollNC.clear();	  
	}   

   if( SUCCEEDED(hr) ) 
     m_bRequiresSave = true;

   return hr;
 }

HRESULT CSafetyPrecaution::InitCollections()
 {
   m_spCollFChange.Release();
   

   
   CComObject<CICollFChange>* pCF;
   HRESULT hr = CComObject<CICollFChange>::CreateInstance( &pCF );
   if( FAILED(hr) )
	{		  
	  PassError( L"CComObject<CICollFChange>::CreateInstance", hr, GetObjectCLSID(), IID_ISafetyPrecaution  );
	  return hr;
	}
   hr = pCF->_InternalQueryInterface( IID_IICollFChange, (void**)&m_spCollFChange );
   if( FAILED(hr) )
	{
	  delete pCF;
	  PassError( L"_InternalQueryInterface of IID_IICollFChange failed", hr, GetObjectCLSID(), IID_ISafetyPrecaution );
	  return hr;
	}
   
   m_vecCollNC.clear();

   return S_OK;
 }


STDMETHODIMP CSafetyPrecaution::get_NonCompatibles(SAFEARRAY** pparrVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pparrVal ) return E_POINTER;

	SAFEARRAYBOUND bnd = { m_vecCollNC.size(), 0 };
	*pparrVal = SafeArrayCreate( VT_I4, 1, &bnd );

	if( !*pparrVal )
	 {
	   Error( L"Cann't array allocate", IID_ISafetyPrecaution, E_FAIL );
	   return E_FAIL;
	 }

	long* pDta;
	HRESULT hr = SafeArrayAccessData( *pparrVal, (void**)&pDta );
	if( FAILED(hr) )
	 {
	   SafeArrayDestroy( *pparrVal );
	   *pparrVal = NULL;	   
	   Error( L"Cann't array access", IID_ISafetyPrecaution, E_FAIL );
	   return E_FAIL;
	 }


    copy( m_vecCollNC.begin(), m_vecCollNC.end(), pDta );
	

	SafeArrayUnaccessData( *pparrVal );


	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::put_NonCompatibles(SAFEARRAY** parrVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !parrVal ) return E_POINTER;

    if( !*parrVal )
	 {
	   m_vecCollNC.clear();	   
	   m_bRequiresSave = true;
	   return S_OK;
	 }

	HRESULT hr = SafeArrayLock( *parrVal );
	if( FAILED(hr) )
	 {	   	   
	   Error( L"Cann't array access", IID_ISafetyPrecaution, E_FAIL );
	   return E_FAIL;
	 }
	if( (*parrVal)->cDims != 1 )
	 {
	   SafeArrayUnlock( *parrVal );
	   Error( L"Array must have 1 dimension only", IID_ISafetyPrecaution, E_FAIL );
	   return E_FAIL;
	 }
	if( (*parrVal)->cbElements != sizeof(long) )
	 {
	   SafeArrayUnlock( *parrVal );
	   Error( L"Array must have 'long' type", IID_ISafetyPrecaution, E_FAIL );
	   return E_FAIL;
	 }

	m_vecCollNC.clear();
	m_vecCollNC.assign( (*parrVal)->rgsabound[0].cElements );
	copy( (long*)(*parrVal)->pvData, ((long*)(*parrVal)->pvData) + (*parrVal)->rgsabound[0].cElements, m_vecCollNC.begin() );

	SafeArrayUnlock( *parrVal );

	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::get_IsDirty(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = (IsDirty() == S_OK ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::get_Profit(CURRENCY *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_ocProfit;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::put_Profit(CURRENCY newVal)
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

STDMETHODIMP CSafetyPrecaution::get_Ke(double *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_dKe;

	return S_OK;
}

STDMETHODIMP CSafetyPrecaution::put_Ke(double newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_dKe = newVal;

	return S_OK;
}


STDMETHODIMP CSafetyPrecaution::Load(LPSTREAM pStm)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Load\n"));

   if( !pStm ) return E_POINTER;
      	  
   CSafetyPrecaution::TInternalData dta;
   HRESULT hr = pStm->Read( &dta, sizeof dta, NULL );
   if( SUCCEEDED(hr) )
	{
	  dta.LoadToObj( *this );

	  hr = m_bsName.ReadFromStream( pStm );
	  if( SUCCEEDED(hr) )
	   {
		 hr = InitCollections();
		 if( FAILED(hr) ) return hr;

		 CComQIPtr<IPersistStreamInit> spPS( m_spCollFChange.p );				
		 hr = spPS->Load( pStm );
		 if( FAILED(hr) ) return hr;	   				

		 long lCntVec;
		 hr = pStm->Read( &lCntVec, sizeof(lCntVec), NULL );

		 if( SUCCEEDED(hr) )
		  {			
			auto_ptr<long> apTmp( new long[lCntVec] );
            hr = pStm->Read( apTmp.get(), lCntVec * sizeof(long), NULL );

			if( SUCCEEDED(hr) )
			 {
			   m_vecCollNC.assign( lCntVec );
			   copy( apTmp.get(), apTmp.get() + lCntVec, m_vecCollNC.begin() );

			   /*vector<long>::iterator it( m_vecCollNC.begin() );
			   for( ; it != m_vecCollNC.end(); ++it )
				{
				  hr = pStm->Read( &*it, sizeof(lCntVec), NULL );
				  if( FAILED(hr) ) break;
				}*/
			 }
		  }

		 if( FAILED(hr) )
		  {
			m_vecCollNC.clear();
			return hr;	   				
		  }
	   }
	}

   
   if( FAILED(hr) )
	 PassError( L"IPersistStreamInit.Save", hr, GetObjectCLSID(), IID_ISafetyPrecaution  );
   else	m_bRequiresSave = false;
	 

   return S_OK;
 }

STDMETHODIMP CSafetyPrecaution::Save(LPSTREAM pStm, BOOL fClearDirty)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Save\n"));

   if( !pStm ) return E_POINTER;
      	   
   CSafetyPrecaution::TInternalData dta;
   dta.LoadFromObj( *this );
   HRESULT hr = pStm->Write( &dta, sizeof dta, NULL );
   if( SUCCEEDED(hr) )
	{
	  hr = m_bsName.WriteToStream( pStm );
	  if( SUCCEEDED(hr) )
	   {
		 CComQIPtr<IPersistStreamInit> spPS( m_spCollFChange.p );				
		 hr = spPS->Save( pStm, fClearDirty );
		 if( FAILED(hr) ) return hr;	   				

		 long lCntVec = m_vecCollNC.size();
		 hr = pStm->Write( &lCntVec, sizeof(lCntVec), NULL );

		 if( SUCCEEDED(hr) )
		  {		
		    auto_ptr<long> apTmp( new long[lCntVec] );
			copy( m_vecCollNC.begin(), m_vecCollNC.end(), apTmp.get() );
			hr = pStm->Write( apTmp.get(), sizeof(long) * lCntVec, NULL );
		    
			/*vector<long>::iterator it( m_vecCollNC.begin() );
			for( ; it != m_vecCollNC.end(); ++it )
			 {
			   hr = pStm->Write( &*it, sizeof(lCntVec), NULL );
			   if( FAILED(hr) ) break;
			 }*/
		  }

		 if( FAILED(hr) )
		  {
			m_vecCollNC.clear();
			return hr;	   				
		  }
	   }
	}
	
   
   if( FAILED(hr) )
	 PassError( L"IPersistStreamInit.Save", hr, GetObjectCLSID(), IID_ISafetyPrecaution  );
   else	if( fClearDirty == TRUE )
	 m_bRequiresSave = false;

   return S_OK;
 }