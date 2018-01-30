#ifndef __OLEENUMITERATOR_H_
#define __OLEENUMITERATOR_H_

#include "resource.h"       // main symbols

template<class T>
struct VARIANTExtractor
 {
//public:
   HRESULT operator()( CComVariant& rVar, T& rResult )
   {         
	 HRESULT hr;
	 if( rVar.vt == VT_DISPATCH )
	   hr = rVar.pdispVal ? rVar.pdispVal->QueryInterface( __uuidof(T), (void**)&rResult ):E_FAIL;
	 else if( rVar.vt == VT_UNKNOWN )
	   hr = rVar.punkVal ? rVar.punkVal->QueryInterface( __uuidof(T), (void**)&rResult ):E_FAIL;
	 else hr = E_FAIL;

	 return hr;
   }
 };
 

template< class COLLECTION, 
  class ENUMINTF, 
  class ITEM, 
  class T, 
  class EXTRACTOR >
class COleEnumIterator
 {
public:
   COleEnumIterator( COLLECTION* pIColl )
	{
	  Init( pIColl );
	}

   HRESULT Init( COLLECTION* pIColl )
	{
	  CComPtr<IUnknown> spNewEnum;
      m_hr = pIColl->get__NewEnum( &spNewEnum );
	  if( SUCCEEDED(m_hr) )
	   {
	     m_spIEnumXX.Release();
         m_hr = spNewEnum.QueryInterface( &m_spIEnumXX );
		 if( SUCCEEDED(m_hr) )
		   m_hr = m_spIEnumXX->Reset();
	   }
	  return m_hr;
	}

   HRESULT Reset()
	{
	  if( m_spIEnumXX.p )
	    m_hr = m_spIEnumXX->Reset();
	  else m_hr = E_FAIL;
	  return m_hr;
	}

   operator HRESULT()
	{
	  return m_hr;
	}

   T operator++( int ) //postfix
	{
	  if( !m_spIEnumXX )
	   {
	     m_hr = E_FAIL; return NULL;
	   }

	  T pRet;
	  ITEM itVar;
	  DWORD lCnt;
	  m_hr = m_spIEnumXX->Next( 1, &itVar, &lCnt );
	  if( m_hr == S_OK )
	   {
	     EXTRACTOR ex;
	     m_hr = ex( itVar, pRet );
	   }

	  return pRet;
	}

   void Release()
	{
	  m_spIEnumXX.Release();
	}

protected:
   HRESULT m_hr;
   CComPtr<ENUMINTF> m_spIEnumXX;
 };

#endif //__COLLFAC_H_
