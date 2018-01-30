// Factor.cpp : Implementation of CFactor
#include "stdafx.h"
#include "GertNet.h"
#include "Factor.h"
#include "LingvoEnum.h"

#include "PassErr.h"

#include <tchar.h>

#include <sstream>
#include <string>
#include <iomanip>
using namespace std;


/////////////////////////////////////////////////////////////////////////////
// CFactor

STDMETHODIMP CFactor::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IFactor,
		&IID_IPersistStreamInit,		
		&IID_IClonable,
		&IID_IBSTRKey
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CFactor::get_Name(BSTR *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pVal == NULL ) return E_POINTER;
	*pVal = m_bsName.Copy();

	return S_OK;
}

STDMETHODIMP CFactor::put_Name(BSTR newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_bsName = newVal;
	m_bRequiresSave = true;
	Fire_OnDirtyChanged( VARIANT_TRUE );

	return S_OK;
}



STDMETHODIMP CFactor::get_IDEnum(long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pVal == NULL ) return E_POINTER;
	*pVal = m_lIDEnum;

	return S_OK;
}

STDMETHODIMP CFactor::put_IDEnum(long lVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	
	m_lIDEnum = lVal;
	m_bRequiresSave = true;
	Fire_OnDirtyChanged( VARIANT_TRUE );

	return S_OK;
}

STDMETHODIMP CFactor::get_Value(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pVal == NULL ) return E_POINTER;
	*pVal = m_sValue;

	return S_OK;
}

STDMETHODIMP CFactor::put_Value(short newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( newVal < 0 || newVal >= CLingvoEnum::GetCount() )
	 {
	   basic_stringstream<WCHAR> strm;
	   strm << L"CFactor::put_Value: value must be: [0 - " << CLingvoEnum::GetCount() - 1 << L"]";
	   PassError( strm.str().c_str(), E_INVALIDARG, GetObjectCLSID(), IID_IFactor );
	   return E_INVALIDARG;  
	 }

	m_sValue = newVal;
	m_bRequiresSave = true;
	Fire_OnDirtyChanged( VARIANT_TRUE );

	return S_OK;
}



// IPersistStreamInit
STDMETHODIMP CFactor::Load( LPSTREAM pStm )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Load\n"));

   /*m_bsName = L"<None>";
   m_bsShortName = L"<None>";
   m_lIDEnum = 0;
   m_sValue = 0;*/

   
   HRESULT hr = m_bsName.ReadFromStream( pStm );
   if( SUCCEEDED(hr) )
	{	     
      hr = m_bsShortName.ReadFromStream( pStm );
      if( SUCCEEDED(hr) )
	   {
		 TInternalData dta;		 
		 hr = pStm->Read( &dta, sizeof dta, NULL );
		 if( SUCCEEDED(hr) ) dta.LoadToObj( *this );
	   }
	}
			
			      
   if( FAILED(hr) )
     PassError( L"IStream.Read", hr, GetObjectCLSID(), IID_ILingvoEnum );   
   else
     m_bRequiresSave = false, Fire_OnDirtyChanged( VARIANT_FALSE );
   
      
   return hr; 
 }
STDMETHODIMP CFactor::Save( LPSTREAM pStm, BOOL fClearDirty )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Save\n"));   

   
   HRESULT hr = m_bsName.WriteToStream( pStm );	     
   if( SUCCEEDED(hr) )
	{
      hr = m_bsShortName.WriteToStream( pStm );
	  if( SUCCEEDED(hr) )
	   {
	     TInternalData dta;
		 dta.LoadFromObj( *this );
		 hr = pStm->Write( &dta, sizeof dta, NULL );
	   }
	}
				
   
   if( FAILED(hr) )
     PassError( L"IStream.Write", hr, GetObjectCLSID(), IID_ILingvoEnum );   
   else
	 m_bRequiresSave = (fClearDirty == TRUE ? false:m_bRequiresSave),
	 Fire_OnDirtyChanged( fClearDirty == TRUE ? VARIANT_FALSE:VARIANT_TRUE );
   
      
   return hr; 
 }

STDMETHODIMP CFactor::InitNew()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::InitNew\n"));
   

   m_bsName = L"<None>";
   m_bsShortName = L"<None>";
   m_lIDEnum = 0;
   m_sValue = 0;
   m_tr = TL_Normal;

   m_pdt = PDT_Uniform;
   m_fPlacement = 0.5; m_fScale = 0.19;

   static float arrshIndStd[ NUMBER_IDX ] = {0,1,-1,-1,-1};
   memcpy( arrshInd, arrshIndStd, sizeof(arrshIndStd) );

   m_bRequiresSave = true;

   return S_OK;
 }

STDMETHODIMP CFactor::get_TrustLvl(TrustLevel *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( pVal == NULL ) return E_POINTER;
	*pVal = m_tr;

	return S_OK;
}

STDMETHODIMP CFactor::put_TrustLvl(TrustLevel newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_tr = newVal;
	m_bRequiresSave = true;
	Fire_OnDirtyChanged( VARIANT_TRUE );

	return S_OK;
}

STDMETHODIMP CFactor::Clone(/*[out, retval]*/ IUnknown** ppUnk)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( ppUnk == NULL ) return E_POINTER;

    CComObject<CFactor>* pObj;
	HRESULT hr = CComObject<CFactor>::CreateInstance( &pObj );
	if( FAILED(hr) )
	 {
       PassError( L"CFactor::Copy: CComObject<CFactor>::CreateInstance", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }

	CComPtr<IUnknown> spTmp;
	hr = pObj->_InternalQueryInterface( IID_IUnknown, (void**)&spTmp );
	if( FAILED(hr) )
	 {
	   delete pObj;
       PassError( L"CFactor::Copy: CComObject<CFactor>::_InternalQueryInterface", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }

	pObj->m_bsName = m_bsName;
	pObj->m_bsShortName = m_bsShortName;
	pObj->m_lIDEnum = m_lIDEnum;
	pObj->m_sValue = m_sValue;
	pObj->m_tr = m_tr;
    memcpy( pObj->arrshInd, arrshInd, sizeof(arrshInd) );
	

	pObj->m_pdt = m_pdt;
    pObj->m_fPlacement = m_fPlacement; 
	pObj->m_fScale = m_fScale;

	pObj->m_bRequiresSave = FALSE;

	spTmp.CopyTo( ppUnk );	

	return S_OK;
 }

STDMETHODIMP CFactor::get_ID(VARIANT *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( !pVal ) return E_POINTER;

	VariantInit( pVal );
	pVal->vt = VT_BSTR;
	V_BSTR(pVal) = m_bsShortName.Copy();

	return S_OK;
}

STDMETHODIMP CFactor::get_IsDirty(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = (IsDirty() == S_OK ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CFactor::ResetOverWrap()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	m_sOverwrap = 0;

	return S_OK;
}

STDMETHODIMP CFactor::AddOwerWrap()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( m_sValue < 0 )
	  m_sOverwrap += -m_sValue;
	else if( m_sValue >= CLingvoEnum::GetCount() )
	  m_sOverwrap += m_sValue - CLingvoEnum::GetCount() + 1;

	return S_OK;
}

STDMETHODIMP CFactor::get_Overwrap(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = m_sOverwrap;

	return S_OK;
}

STDMETHODIMP CFactor::get_Idx(short iIdx, float *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( iIdx >= NUMBER_IDX || iIdx < 0 )
	 {
	   basic_stringstream<WCHAR> strm;
	   strm << L"CFactor::get_Idx: index " << iIdx << L" out of range. Need 0 - " << NUMBER_IDX;
	   Error( strm.str().c_str(), IID_IFactor, E_INVALIDARG );   
       return E_INVALIDARG;
	 }
	*pVal = arrshInd[ iIdx ];

	return S_OK;
}

STDMETHODIMP CFactor::put_Idx(short iIdx, float newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( iIdx >= NUMBER_IDX || iIdx < 0 )
	 {
	   basic_stringstream<WCHAR> strm;
	   strm << L"CFactor::get_Idx: index " << iIdx << L" out of range. Need 0 - " << NUMBER_IDX;
	   Error( strm.str().c_str(), IID_IFactor, E_INVALIDARG );   
       return E_INVALIDARG;
	 }

	if( newVal < 0 )
	 {	   
	   Error( L"Value of index need >= 0", IID_IFactor, E_INVALIDARG );   
       return E_INVALIDARG;
	 }

	arrshInd[ iIdx ] = newVal;
	m_bRequiresSave = true;

	return S_OK;
}

short CFactor::GetNIdx()
 {
   for( int i = 0; i < NUMBER_IDX; ++i )
	  if( arrshInd[i] == -1 ) break;

   return i;
 }

STDMETHODIMP CFactor::get_NIdx(short *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = GetNIdx();

	return S_OK;
}

STDMETHODIMP CFactor::get_DistrType(PrpbDistrTyp *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( pVal == NULL ) return E_POINTER;
	*pVal = m_pdt;

	return S_OK;
}

STDMETHODIMP CFactor::put_DistrType(PrpbDistrTyp newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( newVal != m_pdt )
	 {
	   m_pdt = newVal;
	   m_bRequiresSave = true;
	 }

	return S_OK;
}

STDMETHODIMP CFactor::get_CochiPlacement(float *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pVal == NULL ) return E_POINTER;
	*pVal = m_fPlacement;

	return S_OK;
}

STDMETHODIMP CFactor::put_CochiPlacement(float newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( newVal != m_fPlacement )
	 {
	   m_fPlacement = newVal;
	   m_bRequiresSave = true;
	 }

	return S_OK;
}

STDMETHODIMP CFactor::get_CochiScale(float *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pVal == NULL ) return E_POINTER;
	*pVal = m_fScale;

	return S_OK;
}

STDMETHODIMP CFactor::put_CochiScale(float newVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( newVal != m_fScale )
	 {
	   m_fScale = newVal;
	   m_bRequiresSave = true;
	 }

	return S_OK;
}
