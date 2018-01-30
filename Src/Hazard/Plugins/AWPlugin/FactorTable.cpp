// FactorTable.cpp : Implementation of CFactorTable
#include "stdafx.h"
#include "AWPlugin.h"
#include "FactorTable.h"

/////////////////////////////////////////////////////////////////////////////
// CFactorTable

STDMETHODIMP CFactorTable::InterfaceSupportsErrorInfo(REFIID riid) throw()
{
	static const IID* arr[] = 
	{
		&IID_IFactorTable
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CFactorTable::get_Key(long *pVal) throw()
{
	MGET_PROPERTY( pVal, m_lKey );

	return S_OK;
}


STDMETHODIMP CFactorTable::put_Key(long NewVal) throw()
{
	MPUT_PROPERTY( m_lKey, NewVal );

	return S_OK;
}


STDMETHODIMP CFactorTable::get_Status(ObjStatus *pVal) throw()
{
	MGET_PROPERTY( pVal, m_os );

	return S_OK;
}

HRESULT CFactorTable::CheckArgs( long Number, long IdxStart ) throw()
 {
   HRESULT hr = S_OK;
   try {
	 if( IdxStart != -1 && (m_vec.size() < 0 || IdxStart >= m_vec.size()) )
	   {
		 basic_stringstream<WCHAR> strm;
		 strm << L"Bad value of IdxStart: " << IdxStart;
		 //Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		 AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_INVALIDARG );
		 hr = E_INVALIDARG;
	   }
      else
	  if( Number < 1 && Number != -1 )
	   {
		 basic_stringstream<WCHAR> strm;
		 strm << L"Bad value of Number: " << Number;
		 //Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		 AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_INVALIDARG );
		 hr = E_INVALIDARG;
	   }
	}
   catch( bad_alloc e )
	{	  
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollFTables, A2COLE(e.what()), IID_IFactorTable, hr );
	}

	return hr;
 }

STDMETHODIMP CFactorTable::InsertSlots(long Number, long IdxStart) throw()
{
    HRESULT hr = S_OK;

    try {
	   if( CheckArgs(Number, IdxStart) != S_OK ) hr = E_INVALIDARG;
	   else if( Number == -1 ) hr = E_INVALIDARG;
	   else
		{
		  CFactorSlot cgsTmp( 1 );
		  IT_VEC_CFactorSlot it;
		  if( IdxStart != -1 )
			it = m_vec.begin(), advance( it, IdxStart );
		  else it = m_vec.end();

		  m_vec.insert( it, Number, cgsTmp );

		  Modify();
		  //m_bCalculated = false;
		  UncalcN( -1 );
		}
	 }
	catch( bad_alloc e )
	 {	  
	   hr = E_OUTOFMEMORY;
	 }
	catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollFTables, A2COLE(e.what()), IID_IFactorTable, hr );
	}

	return hr;
}

STDMETHODIMP CFactorTable::RemoveSlots(long Number, long IdxStart) throw()
{
    HRESULT hr = S_OK;

    try {
	   if( CheckArgs(Number, IdxStart) != S_OK ) hr = E_INVALIDARG;
       else if( IdxStart != -1 && m_vec.size() - IdxStart < (Number==-1 ? 0:Number) )
		{
		  basic_stringstream<WCHAR> strm;
		  strm << L"Bad value of Number: " << Number;
		  //Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		  AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_INVALIDARG );
		  hr = E_INVALIDARG;
		}
	   else
        {
		  IT_VEC_CFactorSlot it1, it2;

		  if( IdxStart == -1 )	 
			it1 = m_vec.begin();	  
		  else 
			it1 = m_vec.begin(), advance( it1, IdxStart );

		  if( Number == -1 ) it2 = m_vec.end();
		  else { it2 = it1; advance( it2, Number ); }

		  m_vec.erase( it1, it2 );

		  Modify();
		  //m_bCalculated = false;
		  UncalcN( -1 );
		}
	 }
	catch( bad_alloc e )
	 {	  
	   hr = E_OUTOFMEMORY;
	 }
	catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollFTables, A2COLE(e.what()), IID_IFactorTable, hr );
	}

	return hr;
}

STDMETHODIMP CFactorTable::get_NumberSlots(long *pVal) throw()
{
	*pVal = m_vec.size(); 

	return S_OK;
}

HRESULT CFactorTable::CheckIdx( long lIdx ) throw()
 {
   HRESULT hr = S_OK;
   try {
	 if( lIdx < 0 || lIdx >= m_vec.size() )
	  {
		basic_stringstream<WCHAR> strm;
		strm << L"Bad value of ZeroIndex: " << lIdx;
		//Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_INVALIDARG );
		hr = E_INVALIDARG;
	  }
	}
   catch( bad_alloc e )
	{	  
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollFTables, A2COLE(e.what()), IID_IFactorTable, hr );
	}

   return hr;
 }

STDMETHODIMP CFactorTable::get_NComponentName(long ZeroIndex, BSTR *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	return m_vec[ ZeroIndex ].m_bsComponentName.CopyTo( pVal ); 
}

STDMETHODIMP CFactorTable::put_NComponentName(long ZeroIndex, BSTR newVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	if( m_vec[ ZeroIndex ].m_bsComponentName == newVal ) return S_OK;

    m_vec[ ZeroIndex ].m_bsComponentName = newVal;
	Modify();
	return S_OK; 
}

STDMETHODIMP CFactorTable::get_NLingvoVal(long ZeroIndex, BSTR *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	if( m_vec[ ZeroIndex ].m_bCalculated ||  m_vec[ ZeroIndex ].m_bLocked )
	  return m_vec[ ZeroIndex ].m_bsLingvoVal.CopyTo( pVal ); 
	else
	 {
	   static CComBSTR bsTmp( L"<?>" );
	   return bsTmp.CopyTo( pVal ); 
	 }	  
}

STDMETHODIMP CFactorTable::put_NLingvoVal(long ZeroIndex, BSTR newVal)
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;
	if( !m_vec[ ZeroIndex ].m_bLocked )
	 {
	   AtlReportError( CLSID_CollFTables, L"Нельзя менять оценку для незафиксированного слота", IID_IFactorTable, E_FAIL );
	   return E_FAIL;
	 }

	if( m_vec[ ZeroIndex ].m_bsLingvoVal == newVal ) return S_OK;

    m_vec[ ZeroIndex ].m_bsLingvoVal = newVal;
	Modify();
	return S_OK; 
}

STDMETHODIMP CFactorTable::get_NComment(long ZeroIndex, BSTR *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	return m_vec[ ZeroIndex ].m_bsComment.CopyTo( pVal ); 
}

STDMETHODIMP CFactorTable::put_NComment(long ZeroIndex, BSTR newVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	if( m_vec[ ZeroIndex ].m_bsComment == newVal ) return S_OK;

    m_vec[ ZeroIndex ].m_bsComment = newVal;
	Modify();
	return S_OK; 
}

STDMETHODIMP CFactorTable::get_NWeight(long ZeroIndex, float *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;
	
	MGET_PROPERTY( pVal, m_vec[ ZeroIndex ].m_fWeight );

	return S_OK;
}

STDMETHODIMP CFactorTable::put_NWeight(long ZeroIndex, float newVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	if( newVal < 0 || newVal > 1 )
	 {
       AtlReportError( CLSID_CollFTables, L"Ошибочное значение. Вес должен быть 0 - 1", IID_IFactorTable, E_INVALIDARG );
	   return E_INVALIDARG;
	 }


    MPUT_PROPERTY( m_vec[ ZeroIndex ].m_fWeight, newVal );
	//m_bCalculated = false;
	UncalcN( ZeroIndex );

	return S_OK;
}

STDMETHODIMP CFactorTable::get_NValuation(long ZeroIndex, VARIANT *pfVal) throw()
{	
	HRESULT hr = S_OK;

	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	else if( pfVal != NULL )
	 {
	   VariantInit( pfVal );
	   if( m_vec[ ZeroIndex ].m_bCalculated || m_vec[ ZeroIndex ].m_bLocked )
		 V_VT(pfVal) = VT_R4,
         V_R4(pfVal) = m_vec[ ZeroIndex ].m_fValuation;	   
	 }
	else 
	 {
	   AtlReportError( CLSID_CollFTables, L"Bad parametr 'pfVal' type. Float BYREF is expected", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	
	return hr;		
}

STDMETHODIMP CFactorTable::put_NValuation(long ZeroIndex, VARIANT *pfVal) throw()
{	
	HRESULT hr = S_OK;

	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	else if( !m_vec[ ZeroIndex ].m_bLocked )
	 {
	   AtlReportError( CLSID_CollFTables, L"Чтобы модифицировать оценку, надо заблокировать слот", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	else if( pfVal != NULL )
	 {
	   float fVal;
	   if( !FloatFromVariant(pfVal, fVal) )
		{
		  AtlReportError( CLSID_CollFTables, L"Ошибочный тип параметра: требуется float", IID_IFactorTable, E_INVALIDARG );
	      hr = E_INVALIDARG;
		}
	   else
		{
		  if( fVal < 0 || fVal > 10 )
		   {
		     basic_stringstream<WCHAR> strm;
			 strm << L"Ошибочное значение: " << fVal << L". Нужно: 0 - 10";
			 AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
			 hr = E_INVALIDARG;
		   }
		  else
		   {
			 m_vec[ ZeroIndex ].m_fValuation = fVal;
			 UncalcN( ZeroIndex );
			 m_vec[ ZeroIndex ].Calculate();
			 Modify();		  
		   }
		}
	 }
	else 
	 {
	   AtlReportError( CLSID_CollFTables, L"Bad parametr 'pfVal' type. Float BYREF is expected", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	
	return hr;		
}

/*STDMETHODIMP CFactorTable::put_NValuation(long ZeroIndex, float newVal)
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

    MPUT_PROPERTY_NM( m_vec[ ZeroIndex ].m_fValuation, newVal );

	return S_OK;
}*/

STDMETHODIMP CFactorTable::get_NAveWeighted(long ZeroIndex, VARIANT *pfVal) throw()
{
    HRESULT hr = S_OK;

	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	else if( pfVal != NULL )
	 {
	   VariantInit( pfVal );
	   if( m_vec[ ZeroIndex ].m_bCalculated )
		 V_VT(pfVal) = VT_R4,
         V_R4(pfVal) = m_vec[ ZeroIndex ].m_fAveWeighted;	   
	 }
	else 
	 {
	   AtlReportError( CLSID_CollFTables, L"Bad parametr 'pfVal' type. Float BYREF is expected", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	
	return hr;	
}

/*STDMETHODIMP CFactorTable::put_NAveWeighted(long ZeroIndex, float newVal)
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

    MPUT_PROPERTY_NM( m_vec[ ZeroIndex ].m_fAveWeighted, newVal );

	return S_OK;
}*/

STDMETHODIMP CFactorTable::get_Name(BSTR *pVal) throw()
{	
	return m_bsName.CopyTo( pVal ); 
}

STDMETHODIMP CFactorTable::put_Name(BSTR newVal) throw()
{	
	if( m_bsName == newVal ) return S_OK;

    m_bsName = newVal;
	Modify();
	return S_OK; 
}

STDMETHODIMP CFactorTable::get_Sign(long *pVal) throw()
{
	MGET_PROPERTY( pVal, m_lSign );

	return S_OK;
}

STDMETHODIMP CFactorTable::put_Sign(long NewVal) throw()
{
	MPUT_PROPERTY_NM( m_lSign, NewVal );

	return S_OK;
}


STDMETHODIMP CFactorTable::get_WeightSumm(float *pVal) throw()
{
	MGET_PROPERTY( pVal, m_fWeightSumm );
	return S_OK;
}


STDMETHODIMP CFactorTable::get_ValuationSumm(VARIANT *pfVal) throw()
{
    HRESULT hr = S_OK;
    if( pfVal != NULL )
	 {
	   VariantInit( pfVal );
	   if( m_bCalculated )
         V_VT(pfVal) = VT_R4,
         V_R4(pfVal) = m_fValuationSumm;
	 }
	else 
	 {
	   AtlReportError( CLSID_CollFTables, L"Bad parametr 'pfVal' type. Float BYREF is expected", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	
	return hr;
}


STDMETHODIMP CFactorTable::get_AveWeightedSumm(VARIANT *pfVal) throw()
{	
	HRESULT hr = S_OK;
    if( pfVal != NULL )
	 {
	   VariantInit( pfVal );
	   if( m_bCalculated )
		 V_VT(pfVal) = VT_R4,
         V_R4(pfVal) = m_fAveWeightedSumm;	   
	 }
	else 
	 {
	   AtlReportError( CLSID_CollFTables, L"Bad parametr 'pfVal' type. Float BYREF is expected", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	
	return hr;
}
/*STDMETHODIMP CFactorTable::put_AveWeightedSumm(float newVal)
{
	MPUT_PROPERTY_NM( m_fAveWeightedSumm, newVal );
	return S_OK;
}*/

STDMETHODIMP CFactorTable::get_AllComponentsIsAssigned( VARIANT_BOOL* bRes )  throw()
 {
   CIT_VEC_CFactorSlot it( m_vec.begin() );
   CIT_VEC_CFactorSlot it2( m_vec.end() );

   for( ; it != it2; ++it )
	if( it->IsAllAssigned() == false )
	 {
	   *bRes = VARIANT_FALSE;
       return S_OK;
	 }

   *bRes = VARIANT_TRUE;
   return S_OK;
 }

STDMETHODIMP CFactorTable::SetWeights( VARIANT StartIdx ) throw()
{
    HRESULT hr = S_OK;
	if( m_vec.size() < 1 )
	 {
	   //Error( L"Cann't assign values - the table is empty", IID_IGostTable, E_FAIL );
	   AtlReportError( CLSID_CollFTables, L"Cann't assign values - the table is empty", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	else
	 {
	   int iIdx;

	   VARTYPE vt = (V_VT(&StartIdx) & VT_TYPEMASK);
       if( vt == VT_ERROR ) iIdx = 0;
	   else if( V_ISBYREF(&StartIdx) == VT_BYREF )
		{
		  switch( V_VT(&StartIdx) )
		   {
		     case VT_I2|VT_BYREF:
			   iIdx = *V_I2REF(&StartIdx); break;
		     case VT_I4|VT_BYREF:
			   iIdx = *V_I4REF(&StartIdx); break; 
			 case VT_UI2|VT_BYREF:
			   iIdx = *V_UI2REF(&StartIdx); break;
			 case VT_UI4|VT_BYREF:
			   iIdx = *V_UI4REF(&StartIdx); break;
			 default:
			   AtlReportError( CLSID_CollFTables, L"Bad StartIdx type. Need integer.", IID_IFactorTable, E_INVALIDARG );
	           hr = E_INVALIDARG;
		   };
		}
	   else
		{
		  switch( V_VT(&StartIdx) )
		   {
		     case VT_I2:
			   iIdx = V_I2(&StartIdx); break;
		     case VT_I4:
			   iIdx = V_I4(&StartIdx); break; 
			 case VT_UI2:
			   iIdx = V_UI2(&StartIdx); break;
			 case VT_UI4:
			   iIdx = V_UI4(&StartIdx); break;
			 default:
			   AtlReportError( CLSID_CollFTables, L"Bad StartIdx type. Need integer.", IID_IFactorTable, E_INVALIDARG );
	           hr = E_INVALIDARG;
		   };
		}

	   if( SUCCEEDED(hr) )
		{
		  if( iIdx >= m_vec.size() )
		   {
		     basic_stringstream<WCHAR> strm;
			 strm << L"Inex '" << iIdx << "' out of range. Need 0 - " << m_vec.size() - 1;
		     AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_INVALIDARG );
	         hr = E_INVALIDARG;
		   }
		  else
		   {
			 if( m_vec.size() == 1 )
			   m_vec[ 0 ].m_fWeight = 1;
			 else
			  {
			    float fSumm = 0;
				for( int i = 0; i < iIdx; fSumm += m_vec[ i++ ].m_fWeight );

                float fStep = 1.0f - fSumm;
				if( fStep < 0 ) 
				 {
				   AtlReportError( CLSID_CollFTables, L"Summ of weights exceeds one", IID_IFactorTable, E_FAIL );
	               hr = E_FAIL;
				 }
				else
				 {
				   fStep /= float( m_vec.size() - iIdx );
				   for( i = m_vec.size() - 1; i >= iIdx; --i  )
				     m_vec[ i ].m_fWeight = fStep, m_vec[ i ].Uncalc();
				 }
			  }

			 Modify();
			 //m_bCalculated = false;
			 UncalcN( -1 );
			 hr = S_OK;
		   }
		}
	 }

	return hr;
}

inline float mfabs( float f )
 {
   return f >= 0 ? f:-f;
 }

STDMETHODIMP CFactorTable::get_IsWeightsCorrect( VARIANT* Summ, VARIANT_BOOL *pVal ) throw()
{
   CIT_VEC_CFactorSlot it( m_vec.begin() );
   CIT_VEC_CFactorSlot it2( m_vec.end() );

   float fSumm = 0;
   
   for( ; it != it2; ++it )
	 fSumm += it->m_fWeight;

   if( m_vec.size() != 0 )
     *pVal = (mfabs(fSumm - 1.0f) < 0.0000001f ? VARIANT_TRUE:VARIANT_FALSE);
   else *pVal = VARIANT_TRUE; 

   if( Summ && (V_VT(Summ)&VT_TYPEMASK) != VT_ERROR )
	{
	  VARTYPE vt = V_VT(Summ)&VT_TYPEMASK;

	  bool bRes = true; 
	  if( vt == VT_EMPTY ) V_VT(Summ) = VT_R4, V_R4(Summ) = fSumm;
	  else
	   {
         if( V_ISBYREF(Summ) == VT_BYREF )
		   switch( V_VT(Summ) )
			{
			  case VT_R4|VT_BYREF:
				*V_R4REF(Summ) = fSumm;
				break;
			  case VT_R8|VT_BYREF:
				*V_R8REF(Summ) = fSumm;
				break;			  
			  default:
				bRes = false;
			}
		 else
		   switch( V_VT(Summ) )
			{
			  case VT_R4:
				V_R4(Summ) = fSumm;
				break;
			  case VT_R8:
				V_R8(Summ) = fSumm;
				break;			  
			  default:
				bRes = false;
			}	  
	   }

	  if( !bRes ) return E_INVALIDARG;
	}

   return S_OK;
}



STDMETHODIMP CFactorTable::CalcValues(VARIANT *fValuationSumm, VARIANT *fAveWeightedSumm) throw()
{
    HRESULT hr = S_OK;

	if( m_vec.size() < 1 )
	 {
	   AtlReportError( CLSID_CollFTables, L"There aren't slots to do calculating", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }	
	else
	 {
	   IT_VEC_CFactorSlot it( m_vec.begin() );
	   IT_VEC_CFactorSlot it2( m_vec.end() );

	   m_fWeightSumm = m_fValuationSumm = m_fAveWeightedSumm = 0;
	   for( ; it != it2; ++it )
		{
		  if( !it->IsAllAssigned() )
		   {
		     basic_stringstream<WCHAR> strm;
			 strm << L"The slot " << (long)distance(m_vec.begin(), it) << L" is not bounded";
			 AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_FAIL );
			 hr = E_FAIL;
			 break;
		   }

		  m_fWeightSumm += it->m_fWeight; 

		  
		  
		  if( it->m_ftTyp == FT_Gost || it->m_bLocked )
		   {
		     if( !it->m_bCalculated ) 
			  {	    
			    CFactorSlot::EN_AssErrors enRes;
			    enRes = it->Calculate();
				if( enRes != CFactorSlot::AE_OK )
				 {
				   hr = E_FAIL;
			       break;
				 }
			  }		     
		   }
		  else 
		   {
		     //float fvValuationSumm, fvAveWeightedSumm;
             VARIANT vValuationSumm, vAveWeightedSumm;
			 VariantInit( &vValuationSumm ), VariantInit( &vAveWeightedSumm );
             vAveWeightedSumm.vt = vValuationSumm.vt = VT_R4;
			 //vValuationSumm.pfltVal = &fvValuationSumm;
			 //vAveWeightedSumm.pfltVal = &fvAveWeightedSumm;

			 hr = it->m_Lnk.Self.pFT->CalcValues( &vValuationSumm, &vAveWeightedSumm );
			 if( FAILED(hr) ) break;
			 it->m_fValuation = V_R4(&vValuationSumm);
			 it->m_fAveWeighted = V_R4(&vAveWeightedSumm) * it->m_fWeight;
			 it->m_bCalculated = true;

			 //m_fValuationSumm += it->m_fValuation,
			 //m_fAveWeightedSumm += it->m_fAveWeighted;
		   }

		  m_fValuationSumm += it->m_fValuation,
		  m_fAveWeightedSumm += it->m_fAveWeighted;
		}

	   if( SUCCEEDED(hr) )
		{        
		  m_bCalculated = true;

		  if( !FloatIntoVariant(fValuationSumm, m_fValuationSumm) )
		   {		  
			 AtlReportError( CLSID_CollFTables, L"Bad parametr 'fValuationSumm' type. Float is expected", IID_IFactorTable, E_INVALIDARG );
	         hr = E_INVALIDARG;			
		   }
		  else		  
		  if( !FloatIntoVariant(fAveWeightedSumm, m_fAveWeightedSumm) )
		   {		    
			 AtlReportError( CLSID_CollFTables, L"Bad parametr 'fAveWeightedSumm' type. Float is expected", IID_IFactorTable, E_INVALIDARG );
			 hr = E_INVALIDARG;			
		   }		   
		}
	 }

	return hr;
}

STDMETHODIMP CFactorTable::get_NSlotType(long ZeroIndex, FTSlotTyp *pVal) throw()
{
	HRESULT hr = S_OK;
	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	else
	  *pVal = m_vec[ ZeroIndex ].m_ftTyp;

	return hr;
}

STDMETHODIMP CFactorTable::get_NSlotValue(long ZeroIndex, long *pVal) throw()
{
    HRESULT hr = S_OK;
	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	//else if( !m_vec[ ZeroIndex ].IsAllAssigned() )
	  //*pVal = -1;
	else
	  *pVal = m_vec[ ZeroIndex ].m_Lnk.Gost.iSlot;	

	return hr;
}

STDMETHODIMP CFactorTable::put_NSlotValue(long ZeroIndex, long Value) throw()
{
	HRESULT hr = S_OK;

	try {
	   if( FAILED(CheckIdx(ZeroIndex)) ) 
		 hr = E_INVALIDARG;
	   else
		{
		  CFactorSlot::EN_AssErrors enRes = m_vec[ ZeroIndex ].AssignValue( Value );
		  switch( enRes )
		   {
			 case CFactorSlot::AE_OK:
			   Modify();
			   //m_bCalculated = false;
			   UncalcN( ZeroIndex );
			   m_vec[ ZeroIndex ].Calculate(); 
			   break;
			 case CFactorSlot::AE_SlotUnbounded:
			  {
				basic_stringstream<WCHAR> strm;
				strm << L"Slot " << ZeroIndex << L": '" << Chk(m_vec[ZeroIndex].m_bsComponentName) << L"' " << L"is unbounded";
				AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_FAIL );
				hr = E_FAIL;
				break;
			  }
			 case CFactorSlot::AE_Failed:
			   hr = E_FAIL;
			   break;
			 case CFactorSlot::AE_SlotIdxOutOfRange:
			  {
				basic_stringstream<WCHAR> strm;
				strm << L"Slot " << ZeroIndex << L": '" << Chk(m_vec[ZeroIndex].m_bsComponentName) << L"' - " << L"index " << Value << L" out of range";
				AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_FAIL );
				hr = E_FAIL;
				break;
			  }
			 case CFactorSlot::AE_CanntGetValue:
			  {
				basic_stringstream<WCHAR> strm;
				strm << L"Slot " << ZeroIndex << L": '" << Chk(m_vec[ZeroIndex].m_bsComponentName) << L"' - cann't get value";
				AtlReportError( CLSID_CollFTables, strm.str().c_str(), IID_IFactorTable, E_FAIL );
				hr = E_FAIL;
				break;
			  }
		   };
		}
	 }
	catch( bad_alloc e )
	 {	  
	   hr = E_OUTOFMEMORY;
	 }
	catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollFTables, A2COLE(e.what()), IID_IFactorTable, hr );
	}

	return hr;
}


STDMETHODIMP CFactorTable::NAssSlotType(long ZeroIndex, FTSlotTyp NewTyp, long K1, long K2) throw()
{
	HRESULT hr = S_OK;
	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	else
	 switch( NewTyp )
	  {
		case FT_None:
		  m_vec[ ZeroIndex ].MkNone();
		  break;
		case FT_Self:
		  m_vec[ ZeroIndex ].MkSelf( K1, NULL );
		  break;
		case FT_Gost:
		  m_vec[ ZeroIndex ].MkGost( K1, K2, NULL );
		  break;
	  }

	if( SUCCEEDED(hr) )
	 {
	   //m_bCalculated = false;
	   UncalcN( ZeroIndex );
	   Modify();
	 }

	return hr;
}

STDMETHODIMP CFactorTable::NBindSlot(long ZeroIndex, IDispatch *Object) throw()
{
	HRESULT hr = S_OK;
	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	else	 	   
	  switch( m_vec[ ZeroIndex ].m_ftTyp )
	   {
		 case FT_None:		  
		   AtlReportError( CLSID_CollFTables, L"FT_None slot cann't bind", IID_IFactorTable, E_INVALIDARG );
		   hr = E_INVALIDARG;					  
		   break;

		 case FT_Self:
		  {
		    long lK1 = m_vec[ ZeroIndex ].m_Lnk.Self.lKey;

		    if( Object == NULL ) 
			  m_vec[ ZeroIndex ].MkSelf( lK1, NULL );
			else
			 {
			   CComPtr<IFactorTable> sp;
			   hr = Object->QueryInterface( IID_IFactorTable, (void**)&sp );			
			   if( FAILED(hr) )
				 AtlReportError( CLSID_CollFTables, L"Cann't QI IID_IFactorTable", IID_IFactorTable, hr );
			   else				
			     m_vec[ ZeroIndex ].MkSelf( lK1, sp ),
				 //m_bCalculated = false;
				 UncalcN( ZeroIndex );
			 }
			break;
		  }

		 case FT_Gost:
		   {
			 long lK1 = m_vec[ ZeroIndex ].m_Lnk.Gost.lKeyTopic;
			 long lK2 = m_vec[ ZeroIndex ].m_Lnk.Gost.lKeyGost;

			 if( Object == NULL ) 
			   m_vec[ ZeroIndex ].MkGost( lK1, lK2, NULL );
			 else
			  {
				CComPtr<IGostTable> sp;
				hr = Object->QueryInterface( IID_IGostTable, (void**)&sp );			
				if( FAILED(hr) )
				  AtlReportError( CLSID_CollFTables, L"Cann't QI IID_IGostTable", IID_IFactorTable, hr );
				else				
				 {
				   short iSlot = m_vec[ ZeroIndex ].m_Lnk.Gost.iSlot;				   
				   m_vec[ ZeroIndex ].MkGost( lK1, lK2, sp );
				   m_vec[ ZeroIndex ].m_Lnk.Gost.iSlot = iSlot;

				   //m_bCalculated = false;
				   UncalcN( ZeroIndex );
				 }
			  }
			 break;
		  }
	   }
	 
	return hr;
}

HRESULT CFactorSlot::WriteToStream( LPSTREAM pStm ) throw()
 {
   HRESULT hr = S_OK;

   CFactorSlot::TInternalData dta;
   dta.LoadFromObj( *this );

   hr = pStm->Write( &dta, sizeof dta, NULL );
   if( SUCCEEDED(hr) )
	{
	  hr = m_bsComponentName.WriteToStream( pStm );
	  if( SUCCEEDED(hr) )	   
	   {  
	     hr = m_bsComment.WriteToStream( pStm );	   
		 if( SUCCEEDED(hr) )	   
		    hr = m_bsLingvoVal.WriteToStream( pStm );
	   }

	}

   //if( FAILED(hr) )
	 // ReportStgError( hr, L"WriteToStream", CLSID_CollFTables, IID_IFactorTable );

   return hr;
 }
HRESULT CFactorSlot::ReadFromStream( LPSTREAM pStm ) throw()
 {
   HRESULT hr = S_OK;

   CFactorSlot::TInternalData dta;   

   hr = pStm->Read( &dta, sizeof dta, NULL );
   if( SUCCEEDED(hr) )
	{
	  ReleasePtr();
	  memset( &m_Lnk, 0, sizeof m_Lnk );
	  m_bCalculated = false;
	  dta.LoadToObj( *this );

	  m_bsComponentName.Empty();
	  hr = m_bsComponentName.ReadFromStream( pStm );
	  if( SUCCEEDED(hr) )	  
	   {
	     m_bsComment.Empty();
	     hr = m_bsComment.ReadFromStream( pStm );

		 if( SUCCEEDED(hr) )
		   m_bsLingvoVal.Empty(),
	       hr = m_bsLingvoVal.ReadFromStream( pStm );
	   }
	}

   //if( FAILED(hr) )
	 // ReportStgError( hr, L"ReadFromStream", CLSID_CollFTables, IID_IFactorTable );

   return hr;
 }

STDMETHODIMP CFactorTable::InitNew() throw()
 {         
   m_os = OS_New;
   m_vec.clear();

   m_bsName = L"<Новый>";   
      
   return S_OK;
 }

STDMETHODIMP CFactorTable::Load( LPSTREAM pStm ) throw()
 {   
   HRESULT hr = S_OK;
   
   try {
	 hr = pStm->Read( &m_lKey, sizeof m_lKey, NULL );
	 if( SUCCEEDED(hr) )
	  {        
		 m_bsName.Empty();
		 hr = m_bsName.ReadFromStream( pStm );
		 if( SUCCEEDED(hr) )	   
		  {
			long lTmp; 
			hr = pStm->Read( &lTmp, sizeof lTmp, NULL );
			if( SUCCEEDED(hr) )
			 {
			   CFactorSlot cfsTmp;
			   m_vec.clear();
			   m_vec.assign( lTmp, cfsTmp );
			   for( IT_VEC_CFactorSlot it( m_vec.begin() ); lTmp > 0; --lTmp, ++it )
				 if( FAILED(hr = it->ReadFromStream( pStm ))  ) break;
			 }
		  }	   
	  }

	 m_os = OS_None;
	 //m_bCalculated = false;
	 UncalcN( -1 );
	 //if( FAILED(hr) )
	   //ReportStgError( hr, L"Load", CLSID_CollFTables, IID_IFactorTable );
	}
   catch( bad_alloc e )
	{
	  //USES_CONVERSION;
	  //AtlReportError( CLSID_CollFTables, T2OLE(e.what()), IID_IFactorTable, E_OUTOFMEMORY );
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollFTables, A2COLE(e.what()), IID_IFactorTable, hr );
	}

   return hr;
 }

STDMETHODIMP CFactorTable::Save( LPSTREAM pStm, BOOL fClearDirty ) throw()
 {   
   HRESULT hr = S_OK;
   
   hr = pStm->Write( &m_lKey, sizeof m_lKey, NULL );
   if( SUCCEEDED(hr) )
	{	  	  
	   hr = m_bsName.WriteToStream( pStm );
	   if( SUCCEEDED(hr) )	   
		{
		  long lTmp = m_vec.size();
		  hr = pStm->Write( &lTmp, sizeof lTmp, NULL );
		  if( SUCCEEDED(hr) )
		   {		    
			 for( IT_VEC_CFactorSlot it( m_vec.begin() ); lTmp > 0; --lTmp, ++it )
			   if( FAILED(hr = it->WriteToStream( pStm ))  ) break;
		   }
		}	
	}
   
   
   //if( FAILED(hr) )
	 //ReportStgError( hr, L"Load", CLSID_CollFTables, IID_IFactorTable );
   if( SUCCEEDED(hr) && fClearDirty == TRUE ) m_os = OS_None;

   return hr;
 }

STDMETHODIMP CFactorTable::NSlotSelf(long ZeroIndex, long *Key, IFactorTable **IFT) throw()
{
	HRESULT hr = S_OK;
	if( FAILED(CheckIdx(ZeroIndex)) ) hr = E_INVALIDARG;
	else if( m_vec[ZeroIndex].m_ftTyp != FT_Self )
	 {
	   AtlReportError( CLSID_CollFTables, L"Bad slot type: need Self", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	else
	 {
	   if( Key ) *Key = m_vec[ZeroIndex].m_Lnk.Self.lKey;
	   if( IFT ) 
		{ 
		  *IFT = m_vec[ZeroIndex].m_Lnk.Self.pFT;
		  if( *IFT ) (*IFT)->AddRef();
		}
	 }
	 
	return hr;
}

STDMETHODIMP CFactorTable::NSlotGost(long ZeroIndex, long *KeyTopic, long *KeyGost, IGostTable **IGT) throw()
{
	HRESULT hr = S_OK;
	if( FAILED(CheckIdx(ZeroIndex)) ) hr = E_INVALIDARG;
	else if( m_vec[ZeroIndex].m_ftTyp != FT_Gost )
	 {
	   AtlReportError( CLSID_CollFTables, L"Bad slot type: need Gost", IID_IFactorTable, E_FAIL );
	   hr = E_FAIL;
	 }
	else
	 {
	   if( KeyTopic ) *KeyTopic = m_vec[ZeroIndex].m_Lnk.Gost.lKeyTopic;
	   if( KeyGost ) *KeyGost = m_vec[ZeroIndex].m_Lnk.Gost.lKeyGost;
	   if( IGT ) 
		{ 
		  *IGT = m_vec[ZeroIndex].m_Lnk.Gost.pGT;
		  if( *IGT ) (*IGT)->AddRef();
		}
	 }
	 
	return hr;
}

STDMETHODIMP CFactorTable::HandsOffObjects() throw()
{
    IT_VEC_CFactorSlot it( m_vec.begin() );

	for( ; it != m_vec.end(); ++it )
	  if( it->m_ftTyp == FT_Self )
	   {
	     if( it->m_Lnk.Self.pFT != NULL )
		   it->m_Lnk.Self.pFT->Release(),
		   it->m_Lnk.Self.pFT = NULL;
	   }
	  else if( it->m_ftTyp == FT_Gost )
	   {
	     if( it->m_Lnk.Gost.pGT != NULL )
		   it->m_Lnk.Gost.pGT->Release(),
		   it->m_Lnk.Gost.pGT = NULL;
	   }

	return S_OK;
}


STDMETHODIMP CFactorTable::GetRefToFTable(long Key, VARIANT *RefSlot) throw()
{
    HRESULT hr = S_OK;
	CIT_VEC_CFactorSlot it( m_vec.begin() );
	bool bFl = false;
	for( int i = 0; it != m_vec.end(); ++it, ++i )
	  if( it->m_ftTyp == FT_Self && it->m_Lnk.Self.lKey == Key )
	   {
	     if( !ShortIntoVariant(RefSlot, i) ) 
		  {
		    AtlReportError( CLSID_CollFTables, L"Bad RefSlot type: need one of Integer types", IID_IFactorTable, E_INVALIDARG );
	        hr = E_INVALIDARG;
		  }
		 bFl = true;
		 break;
	   }

	if( !bFl ) VariantInit( RefSlot );

	return hr;
}

STDMETHODIMP CFactorTable::GetRefToGost(long KeyTopic, long KeyGost, VARIANT *RefSlot) throw()
{
	HRESULT hr = S_OK;
	CIT_VEC_CFactorSlot it( m_vec.begin() );
	bool bFl = false;
	for( int i = 0; it != m_vec.end(); ++it, ++i )
	  if( it->m_ftTyp == FT_Gost && it->m_Lnk.Gost.lKeyTopic == KeyTopic && it->m_Lnk.Gost.lKeyGost == KeyGost )
	   {
	     if( !ShortIntoVariant(RefSlot, i) ) 
		  {
		    AtlReportError( CLSID_CollFTables, L"Bad RefSlot type: need one of Integer types", IID_IFactorTable, E_INVALIDARG );
	        hr = E_INVALIDARG;
		  }
		 bFl = true;
		 break;
	   }

	if( !bFl ) VariantInit( RefSlot );

	return hr;
}

STDMETHODIMP CFactorTable::get_Storage( IStorage** Stg ) throw()
 {
   if( !Stg ) return E_POINTER;
   *Stg = NULL;
   return S_OK;
 }

STDMETHODIMP CFactorTable::put_DefaultCU( /*[in]*/VARIANT_BOOL bMode ) throw()
 {   
   m_bDefaultCU = (bMode == VARIANT_TRUE ? true:false);
   return S_OK;
 }

STDMETHODIMP CFactorTable::get_DefaultCU( /*[out,retval]*/VARIANT_BOOL* pbMode ) throw()
 {
   *pbMode = (m_bDefaultCU ? VARIANT_TRUE:VARIANT_FALSE);
   return S_OK;
 }

/*[local]*/BSTR STDMETHODCALLTYPE CFactorTable::NameByRef( void )
 {
   return m_bsName;
 }

STDMETHODIMP CFactorTable::NormalizeWeights() throw()
{    
   IT_VEC_CFactorSlot it( m_vec.begin() );
   IT_VEC_CFactorSlot it2( m_vec.end() );
   
   double dSumm = 0;
   for( ; it != it2; ++it )
	 dSumm += it->m_fWeight;

   double dMul, dAdd;
   if( dSumm != 0 ) dMul = 1.0 / dSumm, dAdd = 0;
   else dMul = 1, dAdd = 1.0 / (m_vec.size() == 0 ? 1:m_vec.size());


   for( it = m_vec.begin(); it != it2; ++it )
	it->m_fWeight = it->m_fWeight * dMul  + dAdd, it->Uncalc(); 

   Modify();
   //m_bCalculated = false;
   UncalcN( -1 );

   return S_OK;
}

STDMETHODIMP CFactorTable::get_NLocked(long ZeroIndex, VARIANT_BOOL *pVal) throw()
{
    HRESULT hr = S_OK;
	if( FAILED(CheckIdx(ZeroIndex)) ) 
	  hr = E_INVALIDARG;
	else
	  *pVal = (m_vec[ ZeroIndex ].m_bLocked == true ? VARIANT_TRUE:VARIANT_FALSE);

	return hr;
}

STDMETHODIMP CFactorTable::put_NLocked(long ZeroIndex, VARIANT_BOOL newVal_) throw()
{
    if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	bool newVal = (newVal_ == VARIANT_TRUE ? true:false);
	if( m_vec[ ZeroIndex ].m_bLocked == newVal ) return S_OK;

    m_vec[ ZeroIndex ].m_bLocked = newVal;
	//m_vec[ ZeroIndex ].m_bCalculated = false;
	UncalcN( ZeroIndex );
	if( newVal )  m_vec[ ZeroIndex ].m_bsComment.Empty();
	m_vec[ ZeroIndex ].Calculate(); 
	Modify();
	return S_OK; 
}
