// FacEditSink.cpp : Implementation of CFacEditSink
#include "stdafx.h"
#include "AnglFont.h"

/*
#include "FacEditSink.h"


/////////////////////////////////////////////////////////////////////////////
// CFacEditSink

STDMETHODIMP CFacEditSink::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IFacEditSink
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}


void CFacEditSink::FinalRelease()
 {   
	RemoveAll();
	if( m_pUnkHelp )
	  m_pUnkHelp->Release(), m_pUnkHelp = NULL;
 }

STDMETHODIMP CFacEditSink::Add(IDispatch *pF)
 {
	if( !pF ) return E_POINTER;

    CComQIPtr<IConnectionPointContainer> spCP( pF );
    if( !spCP )
	 {
       Error( L"No ConnectionPointContainer", IID_IFacEditSink, E_FAIL );
	   return E_FAIL;
	 }
	

	CComPtr<IConnectionPoint> spCPoint;
	HRESULT hr = spCP->FindConnectionPoint( DIID___FacEdit, &spCPoint );
	if( FAILED(hr) ) return hr;

	TCookItem dta; dta.m_dw = 0, dta.m_pDsp = pF; 
    IT_TS_Dsp it = m_set.find( dta );
	if( it != m_set.end() ) 
	 {
       Error( L"This object already advised", IID_IFacEditSink, E_INVALIDARG );
	   return E_INVALIDARG;
	 }

	DWORD dw;
	hr = spCPoint->Advise( m_pUnkHelp, &dw );
    if( FAILED(hr) ) return hr;
	
	m_set.insert( dta );
	pF->AddRef();

	return S_OK;
 }

STDMETHODIMP CFacEditSink::Remove(IDispatch *pF)
 {
	if( !pF ) return E_POINTER;

	TCookItem dta; dta.m_dw = 0, dta.m_pDsp = pF; 
	IT_TS_Dsp it = m_set.find( dta );
	if( it == m_set.end() ) 
	 {
       Error( L"This object is't advised", IID_IFacEditSink, E_INVALIDARG );
	   return E_INVALIDARG;
	 }

	CComQIPtr<IConnectionPointContainer> spCP( pF );
    if( !spCP )
	 {
       Error( L"No ConnectionPointContainer", IID_IFacEditSink, E_FAIL );
	   return E_FAIL;
	 }
	 
	CComPtr<IConnectionPoint> spCPoint;
	HRESULT hr = spCP->FindConnectionPoint( DIID___FacEdit, &spCPoint );
	if( FAILED(hr) ) return hr;

	hr = spCPoint->Unadvise( it->m_dw );

	it->m_pDsp->Release();
	m_set.erase( it );

	return hr;
 }

STDMETHODIMP CSnkHelp::ClickAssEnum(_FacEdit* feSender)
 { 
   m_pS->Fire_ClickAssEnum( feSender );
   return S_OK;
 }
STDMETHODIMP CSnkHelp::ClickSetupValue(_FacEdit* feSender)
 {
   m_pS->Fire_ClickSetupValue( feSender );
   return S_OK;
 }
STDMETHODIMP CSnkHelp::FactorChanged(IFactor* fac)
 {
   m_pS->Fire_FactorChanged( fac ); 
   return S_OK;
 }

STDMETHODIMP CFacEditSink::RemoveAll(void)
{
	IT_TS_Dsp it( m_set.begin() );
	for( ; it != m_set.end(); ++it )
	 {
       CComQIPtr<IConnectionPointContainer> spCP( it->m_pDsp );
	   if( !spCP ) continue;
				
	   CComPtr<IConnectionPoint> spCPoint;
	   HRESULT hr = spCP->FindConnectionPoint( DIID___FacEdit, &spCPoint );
	   if( FAILED(hr) ) continue;

	   hr = spCPoint->Unadvise( it->m_dw );
	   it->m_pDsp = NULL;
	 }

	return S_OK;
}
*/