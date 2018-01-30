// CollFTables.cpp : Implementation of CCollFTables
#include "stdafx.h"
#include "AWPlugin.h"
#include "FactorTable.h"
#include "CollFTables.h"




/////////////////////////////////////////////////////////////////////////////
// CCollFTables

STDMETHODIMP CCollFTables::InterfaceSupportsErrorInfo(REFIID riid) throw()
{
	static const IID* arr[] = 
	{
		&IID_ICollFTables
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CCollFTables::DetectRoot( /*[out,retval]*/long* pKey ) throw()
 {
   if( !pKey ) return E_POINTER;

   HRESULT hr;

   try {
	  set<long> setLong;
	  CIT_TMAP_DIR it( m_mapDir.begin() );
	  for( ; it != m_mapDir.end(); ++it )  setLong.insert( it->second );

	  CComPtr<IFactorTable> sp;
	  VARIANT vtKey;   

	  for( it = m_mapDir.begin(); it != m_mapDir.end(); ++it )
	   {
		 sp.Release(); 

		 V_VT(&vtKey) = VT_I4;
		 V_I4(&vtKey) = it->second;

		 hr = Item( &vtKey, &sp );
		 if( FAILED(hr) ) return hr;
		 long lNSlots; sp->get_NumberSlots( &lNSlots );
		 for( --lNSlots; lNSlots >= 0; --lNSlots )
		  {
			FTSlotTyp ftt;
			hr = sp->get_NSlotType( lNSlots, &ftt );
			if( FAILED(hr) ) return hr;
			if( ftt != FT_Self ) continue;

			long lKTmp;
			hr = sp->NSlotSelf( lNSlots, &lKTmp, NULL );
			if( FAILED(hr) ) return hr;
			
			setLong.erase( lKTmp );
		  }	  
	   }

	  if( setLong.size() != 1 )
	   {
		 basic_stringstream<WCHAR> strm;
		 strm << L"Ошибка поиска корня иерархии: коллекция " << m_lKey << L"; невязка = " << setLong.size();
		 hr = E_FAIL;
		 Error( strm.str().c_str(), IID_ICollFTables, hr );
		 return hr;
	   }
	  else *pKey = *setLong.begin();
	}
   catch( bad_alloc e )
	{	  
	  return E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollFTables, A2COLE(e.what()), IID_ICollFTables, hr );
	  return hr;
	}

   return S_OK;
 }
