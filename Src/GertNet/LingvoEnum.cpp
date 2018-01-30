// LingvoEnum.cpp : Implementation of CLingvoEnum
#include "stdafx.h"
#include "GertNet.h"
#include "LingvoEnum.h"
#include "PassErr.h"

#include <tchar.h>

#include <sstream>
#include <string>
#include <iomanip>
using namespace std;


/////////////////////////////////////////////////////////////////////////////
// CLingvoEnum

LingvoDescr CLingvoEnum::arrLD_Default[ NUMBER_LD ] =
 {
	{L"Очень очень низкое",  0.084},
	{L"Очень низкое",        0.166},
	{L"Низкое",              0.250},
	{L"Ниже среднего",       0.334},
	{L"Среднее",             0.417},
	{L"Выше среднего",       0.500},
	{L"Хорошее",             0.584},
	{L"Очень хорошее",       0.667},
	{L"Высокое",             0.750},
	{L"Очень высокое",       0.834},
	{L"Очень очень высокое", 0.916}
 };

STDMETHODIMP CLingvoEnum::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ILingvoEnum,
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

// ILingvoEnum
void CLingvoEnum::Set_OrderOutRange()
 {
   basic_stringstream<WCHAR> strm;
   strm << L"Order should be in range 0 - " << NUMBER_LD - 1;
   Error( strm.str().c_str(), IID_ILingvoEnum, E_FAIL );   
 }

STDMETHODIMP CLingvoEnum::get_Mark( /*[in]*/ long lOrder, /*[out, retval]*/ BSTR *pVal )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( pVal == NULL ) return E_POINTER;
   if( lOrder < 0 || lOrder >= NUMBER_LD ) 
	{
	  Set_OrderOutRange();	  
	  return E_FAIL;
	}

   *pVal = arrLD[ lOrder ].m_cbsMark.Copy(); 

   return S_OK;
 }
STDMETHODIMP CLingvoEnum::put_Mark( /*[in]*/ long lOrder, /*[in]*/ BSTR newVal )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( lOrder < 0 || lOrder >= NUMBER_LD ) 
	{
	  Set_OrderOutRange();	  
	  return E_FAIL;
	}

   arrLD[ lOrder ].m_cbsMark = newVal;
   SetDirty( TRUE );

   return S_OK;

 }
STDMETHODIMP CLingvoEnum::get_Quality( /*[in]*/ long lOrder, /*[out, retval]*/ double *pVal )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( pVal == NULL ) return E_POINTER;
   if( lOrder < 0 || lOrder >= NUMBER_LD ) 
	{
	  Set_OrderOutRange();	  
	  return E_FAIL;
	}

   *pVal = arrLD[ lOrder ].m_dQuality;

   return S_OK;
 }
STDMETHODIMP CLingvoEnum::put_Quality( /*[in]*/ long lOrder, /*[in]*/ double newVal )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( lOrder < 0 || lOrder >= NUMBER_LD ) 
	{
	  Set_OrderOutRange();	  
	  return E_FAIL;
	}

   arrLD[ lOrder ].m_dQuality = newVal;
   SetDirty( TRUE );

   return S_OK;
 }


// IPersistStreamInit
STDMETHODIMP CLingvoEnum::Load( LPSTREAM pStm )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Load\n"));

   DWORD dwCnt;
   DWORD dwarrTmp[ 2 ];
   HRESULT hr = pStm->Read( &dwarrTmp, sizeof(dwarrTmp), NULL );
   if( FAILED(hr) )
	{
	  PassError( L"IStream.Write", hr, GetObjectCLSID(), IID_ILingvoEnum );   
	  return hr;
	}

   dwCnt = dwarrTmp[ 0 ];
   m_lID = dwarrTmp[ 1 ];

   if( dwCnt != NUMBER_LD )
	{      
      Error( L"Incomatible version of enumerator", IID_ILingvoEnum, E_FAIL );   
	  return E_FAIL;
	}

   for( int i = 0; i < NUMBER_LD; ++i )
	{
      hr = pStm->Read( &arrLD[i].m_dQuality, sizeof(arrLD[i].m_dQuality), NULL );
	  if( FAILED(hr) )
	   {
	     PassError( L"IStream.Read", hr, GetObjectCLSID(), IID_ILingvoEnum  );   
	     return hr;
	   }

	  hr = arrLD[i].m_cbsMark.ReadFromStream( pStm ); 
	  if( FAILED(hr) )
	   {
	     PassError( L"IStream.Read", hr, GetObjectCLSID(), IID_ILingvoEnum  );   
	     return hr;
	   }
	}

   SetDirty( FALSE );

   return S_OK; 
 }
STDMETHODIMP CLingvoEnum::Save( LPSTREAM pStm, BOOL fClearDirty )
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::Save\n"));   

   DWORD dwarrTmp[ 2 ] = { NUMBER_LD, m_lID };
   //DWORD dwCnt = NUMBER_LD;
   HRESULT hr = pStm->Write( &dwarrTmp, sizeof(dwarrTmp), NULL );
   if( FAILED(hr) )
	{
	  PassError( L"IStream.Write", hr, GetObjectCLSID(), IID_ILingvoEnum  );   
	  return hr;
	}

   for( int i = 0; i < NUMBER_LD; ++i )
	{
      hr = pStm->Write( &arrLD[i].m_dQuality, sizeof(arrLD[i].m_dQuality), NULL );
	  if( FAILED(hr) )
	   {
	     PassError( L"IStream.Write", hr, GetObjectCLSID(), IID_ILingvoEnum  );   
	     return hr;
	   }

	  hr = arrLD[i].m_cbsMark.WriteToStream( pStm ); 
	  if( FAILED(hr) )
	   {
	     PassError( L"IStream.Write", hr, GetObjectCLSID(), IID_ILingvoEnum  );   
	     return hr;
	   }
	}

   if( fClearDirty == TRUE ) SetDirty( FALSE );

   return S_OK;
 }

STDMETHODIMP CLingvoEnum::InitNew()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   ATLTRACE2(atlTraceCOM, 0, _T("IPersistStreamInit::InitNew\n"));

   m_lID = 0;

   for( int i = sizeof(CLingvoEnum::arrLD_Default) / sizeof(CLingvoEnum::arrLD_Default[0]) - 1; i >= 0; --i )
	  arrLD[ i ].m_dQuality = CLingvoEnum::arrLD_Default[ i ].m_dQuality,
	  arrLD[ i ].m_cbsMark = CLingvoEnum::arrLD_Default[ i ].m_bsMark;

   SetDirty( TRUE );

   return S_OK;
 }




STDMETHODIMP CLingvoEnum::get_Count(long *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pVal == NULL ) return E_POINTER;

	*pVal = NUMBER_LD;

	return S_OK;
}


STDMETHODIMP CLingvoEnum::Clone(/*[out, retval]*/ IUnknown** ppUnk)
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( ppUnk == NULL ) return E_POINTER;

    CComObject<CLingvoEnum>* pObj;
	HRESULT hr = CComObject<CLingvoEnum>::CreateInstance( &pObj );
	if( FAILED(hr) )
	 {
       PassError( L"CLingvoEnum::Copy: CComObject<CLingvoEnum>::CreateInstance", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }

	CComPtr<IUnknown> spTmp;
	hr = pObj->_InternalQueryInterface( IID_IUnknown, (void**)&spTmp );
	if( FAILED(hr) )
	 {
	   delete pObj;
       PassError( L"CLingvoEnum::Copy: CComObject<CLingvoEnum>::_InternalQueryInterface", hr, GetObjectCLSID(), IID_IMGertNet );
       return hr;
	 }

	pObj->m_lID = m_lID;
	for( int i = sizeof(arrLD) / sizeof(arrLD[0]) - 1; i >= 0; --i )
	  pObj->arrLD[ i ].m_dQuality = arrLD[ i ].m_dQuality,
	  pObj->arrLD[ i ].m_cbsMark = arrLD[ i ].m_cbsMark;

	pObj->m_bRequiresSave = FALSE;

	
	spTmp.CopyTo( ppUnk );

	return S_OK;
 }

STDMETHODIMP CLingvoEnum::MarkForValue(double dVal, BSTR *pbsMark)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( pbsMark == NULL ) return E_POINTER;

	for( int i = sizeof(arrLD) / sizeof(arrLD[0]) - 1; i >= 0; --i )
	 if( arrLD[i].m_dQuality == dVal )
	  {
	    *pbsMark = arrLD[i].m_cbsMark.Copy();
		return S_OK;
	  }

   basic_stringstream<WCHAR> strm;
   strm << L"Bad value of factor: " << dVal;
   Error( strm.str().c_str(), IID_ILingvoEnum, E_INVALIDARG );   

   return E_INVALIDARG;
}

STDMETHODIMP CLingvoEnum::get_ID(VARIANT *pVal)
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

STDMETHODIMP CLingvoEnum::get_IsDirty(VARIANT_BOOL *pVal)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	*pVal = (IsDirty() == S_OK ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CLingvoEnum::UpdateFrom(ILingvoEnum *pLEnum)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pLEnum ) return E_POINTER;

    CComObject<CLingvoEnum>* pSrc = static_cast<CComObject<CLingvoEnum>*>( pLEnum );

	for( int i = sizeof(arrLD) / sizeof(arrLD[0]) - 1; i >= 0; --i )
	  arrLD[ i ].m_dQuality = pSrc->arrLD[ i ].m_dQuality,
	  arrLD[ i ].m_cbsMark = pSrc->arrLD[ i ].m_cbsMark;

	m_bRequiresSave = TRUE;

	return S_OK;
}

STDMETHODIMP CLingvoEnum::RoundS(double Val, double *dV)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( dV == NULL ) return E_POINTER;

   double dMinDelta = 10;
   short shMinIdx = -1;
   for( int i = 0; i < sizeof(arrLD) / sizeof(arrLD[0]); ++i )
	{
	  double dDelta = fabs( arrLD[i].m_dQuality  - Val );
	  if( dDelta < dMinDelta ) dMinDelta = dDelta, shMinIdx = i;
	  else break;
	}

   *dV = arrLD[ shMinIdx ].m_dQuality;

   return S_OK;
}

STDMETHODIMP CLingvoEnum::ValueIdx(double Val, short *shIdx)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


   if( shIdx == NULL ) return E_POINTER;

   for( int i = sizeof(arrLD) / sizeof(arrLD[0]) - 1; i >= 0; --i )
	if( arrLD[i].m_dQuality == Val )
	 {
	   *shIdx = i;
	   return S_OK;
	 }

   basic_stringstream<WCHAR> strm;
   strm << L"Bad value of factor: " << Val;
   Error( strm.str().c_str(), IID_ILingvoEnum, E_INVALIDARG );   

   return E_INVALIDARG;
}
