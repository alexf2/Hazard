// FChange.cpp : Implementation of CFChange
#include "stdafx.h"
#include "GertNet.h"
#include "FChange.h"

#include "PassErr.h"

#include <tchar.h>

#include <sstream>
#include <string>
#include <iomanip>
using namespace std;


/////////////////////////////////////////////////////////////////////////////
// CFChange

STDMETHODIMP CFChange::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IFChange,
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

STDMETHODIMP CFChange::get_NameFactor(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	return m_bsName.CopyTo(pVal);
}

STDMETHODIMP CFChange::put_NameFactor(BSTR newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_bsName = newVal;
	m_bRequiresSave = true;

	return S_OK;
}

STDMETHODIMP CFChange::get_Delta(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pVal ) return E_POINTER;
	*pVal = m_shDelta;

	return S_OK;
}

STDMETHODIMP CFChange::put_Delta(short newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_shDelta = newVal;
	m_bRequiresSave = true;

	return S_OK;
}


STDMETHODIMP CFChange::Clone(/*[out, retval]*/ IUnknown** ppUnk)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( ppUnk == NULL ) return E_POINTER;

    CComObject<CFChange>* pObj;
	HRESULT hr = CComObject<CFChange>::CreateInstance( &pObj );
	if( FAILED(hr) )
	 {
       PassError( L"CFChange::Copy: CComObject<CFChange>::CreateInstance", hr, GetObjectCLSID(), IID_IFChange );
       return hr;
	 }

	CComPtr<IUnknown> spTmp;
	hr = pObj->_InternalQueryInterface( IID_IUnknown, (void**)&spTmp );
	if( FAILED(hr) )
	 {
	   delete pObj;
       PassError( L"CFChange::Copy: CComObject<CFChange>::_InternalQueryInterface", hr, GetObjectCLSID(), IID_IFChange );
       return hr;
	 }

	pObj->m_lID = m_lID;

	pObj->m_shDelta = m_shDelta;
	pObj->m_bsName = m_bsName;
	pObj->m_tcTrustChange = m_tcTrustChange;
		

	pObj->m_bRequiresSave = false;

	
	spTmp.CopyTo( ppUnk );

	return S_OK;
 }

// IPersistStreamInit
STDMETHODIMP CFChange::Load( LPSTREAM pStm )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Load\n"));

   
   HRESULT hr = m_bsName.ReadFromStream( pStm );
   if( SUCCEEDED(hr) )
	{
	  CFChange::TInternalData dta;
	  hr = pStm->Read( &dta, sizeof dta, NULL );
      if( SUCCEEDED(hr) ) dta.LoadToObj( *this );
	}

   if( FAILED(hr) )
	{
	  PassError( L"IStream.Read: ", hr, GetObjectCLSID(), IID_IFChange  );   
	  return hr;
	}

   m_bRequiresSave = false;

   return S_OK; 
 }
STDMETHODIMP CFChange::Save( LPSTREAM pStm, BOOL fClearDirty )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Save\n"));   

   
   
   HRESULT hr = m_bsName.WriteToStream( pStm );
   if( SUCCEEDED(hr) )
	{
      CFChange::TInternalData dta;
	  dta.LoadFromObj( *this );
	  hr = pStm->Write( &dta, sizeof(dta), NULL );
	}
		 

   if( FAILED(hr) )
	{
	  PassError( L"IStream.Write: ", hr, GetObjectCLSID(), IID_IFChange  );   
	  return hr;
	}

   
   if( fClearDirty == TRUE ) m_bRequiresSave = false;

   return S_OK;
 }


STDMETHODIMP CFChange::InitNew()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::InitNew\n"));

   m_lID = 0;

   m_shDelta = 0;
   m_tcTrustChange = TR_NoChange;
   m_bsName = L"<Not assigned>";

   m_bRequiresSave = true;

   return S_OK;
 }

STDMETHODIMP CFChange::get_ID(VARIANT *pVal)
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

STDMETHODIMP CFChange::get_TCh(TrustChange *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( !pVal ) return E_POINTER;
	*pVal = m_tcTrustChange;

	return S_OK;
}

STDMETHODIMP CFChange::put_TCh(TrustChange newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_tcTrustChange = newVal;
	m_bRequiresSave = true;

	return S_OK;
}
