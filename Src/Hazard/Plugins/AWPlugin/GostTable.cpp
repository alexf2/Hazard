// GostTable.cpp : Implementation of CGostTable
#include "stdafx.h"
#include "AWPlugin.h"
#include "GostTable.h"

/////////////////////////////////////////////////////////////////////////////
// CGostTable

STDMETHODIMP CGostTable::InterfaceSupportsErrorInfo(REFIID riid) throw()
{
	static const IID* arr[] = 
	{
		&IID_IGostTable
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

HRESULT CGostTable::FinalConstruct() throw()
 {   
   m_lKey = -1;
   m_os = OS_None;
   m_lSign = 0;
   m_shLastErr = -1;
   m_bDefaultCU = true;
   return S_OK;
 } 

STDMETHODIMP CGostTable::get_Status(ObjStatus *pVal) throw()
{	
	MGET_PROPERTY( pVal, m_os );

	return S_OK;
}

STDMETHODIMP CGostTable::get_Key(long *pVal) throw()
{
	MGET_PROPERTY( pVal, m_lKey );

	return S_OK;
}

STDMETHODIMP CGostTable::put_Key(long NewVal) throw()
{
	MPUT_PROPERTY( m_lKey, NewVal );

	return S_OK;
}



STDMETHODIMP CGostTable::get_Sign(long *pVal)  throw()
{
	MGET_PROPERTY( pVal, m_lSign );

	return S_OK;
}

STDMETHODIMP CGostTable::put_Sign(long NewVal) throw()
{
	MPUT_PROPERTY_NM( m_lSign, NewVal );

	return S_OK;
}


STDMETHODIMP CGostTable::get_NumberSlots(long *pVal) throw()
{
	*pVal = m_vec.size(); 

	return S_OK;
}


HRESULT CGostTable::CheckArgs( long Number, long IdxStart ) throw()
 {
   HRESULT hr = S_OK;
   try {
	  if( IdxStart != -1 && (m_vec.size() < 0 || IdxStart >= m_vec.size()) )
		{
		  basic_stringstream<WCHAR> strm;
		  strm << L"Bad value of IdxStart: " << IdxStart;
		  //Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		  AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		  hr = E_INVALIDARG;
		}
       else
	   if( Number < 1 && Number != -1 )
		{
		  basic_stringstream<WCHAR> strm;
		  strm << L"Bad value of Number: " << Number;
		  //Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		  AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
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
	  AtlReportError( CLSID_CollTopics, A2COLE(e.what()), IID_IGostTable, hr );
	}

	return hr;
 }

STDMETHODIMP CGostTable::InsertSlots(long Number, long IdxStart) throw()
{
    HRESULT hr = S_OK;
    try {
	   if( Number < 1 || 
		   IdxStart < -1 || 
		   (IdxStart > m_vec.size() && IdxStart != -1) ) hr = E_INVALIDARG;
	   else
		{
		  CGostSlot cgsTmp( 1 );
		  IT_VEC_CGostSlot it;
		  if( IdxStart != -1 )
			it = m_vec.begin(), advance( it, IdxStart );
		  else it = m_vec.end();

		  m_vec.insert( it, Number, cgsTmp );

		  Modify();
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
	  AtlReportError( CLSID_CollTopics, A2COLE(e.what()), IID_IGostTable, hr );
	}

	return hr;
}

STDMETHODIMP CGostTable::RemoveSlots(long Number, long IdxStart) throw()
{
    HRESULT hr = S_OK;
    try {
	   if( CheckArgs(Number, IdxStart) != S_OK ) hr = E_INVALIDARG;
       else
	   if( IdxStart != -1 && m_vec.size() - IdxStart < (Number==-1 ? 0:Number) )
		{
		  basic_stringstream<WCHAR> strm;
		  strm << L"Bad value of Number: " << Number;
		  //Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		  AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
		  hr = E_INVALIDARG;
		}
	   else
		{
		  IT_VEC_CGostSlot it1, it2;

		  if( IdxStart == -1 )	 
			it1 = m_vec.begin();	  
		  else 
			it1 = m_vec.begin(), advance( it1, IdxStart );

		  if( Number == -1 ) it2 = m_vec.end();
		  else { it2 = it1; advance( it2, Number ); }

		  m_vec.erase( it1, it2 );

		  Modify();
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
	  AtlReportError( CLSID_CollTopics, A2COLE(e.what()), IID_IGostTable, hr );
	}

	return hr;
}

HRESULT CGostSlot::WriteToStream( LPSTREAM lpStrm ) throw()
 {
   ATLTRACE2(atlTraceCOM, 0, _T("CGostSlot::WriteToStream\n"));   

   HRESULT hr = m_bsValDescr.WriteToStream( lpStrm );	     
   if( SUCCEEDED(hr) )
	{
      hr = m_bsLingvoVal.WriteToStream( lpStrm );
	  if( SUCCEEDED(hr) )
	   {
	     hr = m_bsComment.WriteToStream( lpStrm );
         if( SUCCEEDED(hr) )  
		  {
		    hr = lpStrm->Write( &m_fVal, sizeof m_fVal, NULL );			
		  }
	   }
	}

   //if( FAILED(hr) ) 
	  //return ReportStgError( hr, L"CGostSlot::WriteToStream", CLSID_CollTopics, IID_IGostTable );

   return hr;
 }
HRESULT CGostSlot::ReadFromStream( LPSTREAM lpStrm ) throw()
 {
   ATLTRACE2(atlTraceCOM, 0, _T("CGostSlot::ReadFromStream\n"));   

   m_bsValDescr.Empty();
   HRESULT hr = m_bsValDescr.ReadFromStream( lpStrm );	     
   if( SUCCEEDED(hr) )
	{
	  m_bsLingvoVal.Empty();
      hr = m_bsLingvoVal.ReadFromStream( lpStrm );
	  if( SUCCEEDED(hr) )
	   {
	     m_bsComment.Empty();
	     hr = m_bsComment.ReadFromStream( lpStrm );
         if( SUCCEEDED(hr) )  
		  {
		    hr = lpStrm->Read( &m_fVal, sizeof m_fVal, NULL );			
		  }
		     
	   }
	}
   

   //if( FAILED(hr) ) 
	 // return ReportStgError( hr, L"CGostSlot::ReadFromStream", CLSID_CollTopics, IID_IGostTable );

   return hr;
 }

HRESULT CGostTable::CheckIdx( long lIdx ) throw()
 {
   if( lIdx < 0 || lIdx >= m_vec.size() )
	 {
	   basic_stringstream<WCHAR> strm;
	   strm << L"Bad value of ZeroIndex: " << lIdx;
	   //Error( strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
	   AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
	   return E_INVALIDARG;
	 }
   return S_OK;
 }

STDMETHODIMP CGostTable::get_NDescr(long ZeroIndex, BSTR *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	return m_vec[ ZeroIndex ].m_bsValDescr.CopyTo( pVal ); 
}

STDMETHODIMP CGostTable::put_NDescr(long ZeroIndex, BSTR newVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	if( m_vec[ ZeroIndex ].m_bsValDescr == newVal ) return S_OK;

    m_vec[ ZeroIndex ].m_bsValDescr = newVal;
	Modify();
	return S_OK; 
}

STDMETHODIMP CGostTable::get_NLingvoVal(long ZeroIndex, BSTR *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	return m_vec[ ZeroIndex ].m_bsLingvoVal.CopyTo( pVal ); 
}

STDMETHODIMP CGostTable::put_NLingvoVal(long ZeroIndex, BSTR newVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	if( m_vec[ ZeroIndex ].m_bsLingvoVal == newVal ) return S_OK;

    m_vec[ ZeroIndex ].m_bsLingvoVal = newVal;
	Modify();
	return S_OK; 
}

STDMETHODIMP CGostTable::get_NComment(long ZeroIndex, BSTR *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	return m_vec[ ZeroIndex ].m_bsComment.CopyTo( pVal ); 
}

STDMETHODIMP CGostTable::put_NComment(long ZeroIndex, BSTR newVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	if( m_vec[ ZeroIndex ].m_bsComment == newVal ) return S_OK; 

    m_vec[ ZeroIndex ].m_bsComment = newVal;
	Modify();
	return S_OK; 
}

STDMETHODIMP CGostTable::get_NValue(long ZeroIndex, float *pVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;

	*pVal = m_vec[ ZeroIndex ].m_fVal; 
	return S_OK;
}

STDMETHODIMP CGostTable::put_NValue(long ZeroIndex, float newVal) throw()
{
	if( FAILED(CheckIdx(ZeroIndex)) ) return E_INVALIDARG;
	if( newVal < 0 || newVal > 10 )
	 {
	   basic_stringstream<WCHAR> strm;
	   strm << L"Bad value: " << newVal << L". Need: 0 - 10";
	   AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, E_INVALIDARG );
	   return E_INVALIDARG;
	 }
    
	MPUT_PROPERTY( m_vec[ ZeroIndex ].m_fVal, newVal );
	return S_OK; 
}

STDMETHODIMP CGostTable::get_Name(BSTR *pVal) throw()
{	
	return m_bsName.CopyTo( pVal );
}

STDMETHODIMP CGostTable::put_Name(BSTR newVal) throw()
{
	if( m_bsName == newVal ) return S_OK;

	m_bsName = newVal;
	Modify();
	return S_OK;
}

STDMETHODIMP CGostTable::InitNew() throw()
 {   
   m_bsName = L"<Новый>";
   m_bsValClmName = L"<Нет>";
   m_vec.clear();

   m_os = OS_New;   

   return S_OK;
 }

STDMETHODIMP CGostTable::Load( LPSTREAM pStm ) throw()
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
		   m_bsValClmName.Empty();
		   hr = m_bsValClmName.ReadFromStream( pStm );
		   if( SUCCEEDED(hr) )
			{
			  long lN;
			  hr = pStm->Read( &lN, sizeof lN, NULL );
			  if( SUCCEEDED(hr) )
			   {
				 m_vec.assign( lN );
				 for( int i = 0; lN > 0; --lN, ++i )
				   if( FAILED(hr = m_vec[i].ReadFromStream(pStm)) )
					 break;

				 if( FAILED(hr) ) return hr;
			   }
			}
		 }	   
	  }

	 //if( FAILED(hr) ) 
	//	return ReportStgError( hr, L"CGostTable::Load", CLSID_CollTopics, IID_IGostTable );
	  
	 m_os = OS_None;
	}
   catch( bad_alloc e )
	{	  
	  hr = E_OUTOFMEMORY;
	}
   catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( CLSID_CollTopics, A2COLE(e.what()), IID_IGostTable, hr );
	}

   return hr;
 }

STDMETHODIMP CGostTable::Save( LPSTREAM pStm, BOOL fClearDirty ) throw()
 {
   HRESULT hr = pStm->Write( &m_lKey, sizeof m_lKey, NULL );
   if( SUCCEEDED(hr) )
	{	  
	  hr = m_bsName.WriteToStream( pStm );
	  if( SUCCEEDED(hr) )
	   {
		 hr = m_bsValClmName.WriteToStream( pStm );
		 if( SUCCEEDED(hr) )
		  {
			long lN = m_vec.size();
			hr = pStm->Write( &lN, sizeof lN, NULL );
            if( SUCCEEDED(hr) )
			 {				  
			   for( int i = 0; lN > 0; --lN, ++i )
				 if( FAILED(hr = m_vec[i].WriteToStream(pStm)) )
				   break;

			   if( FAILED(hr) ) return hr;
			 }
		  }
	   }	
	}

   //if( FAILED(hr) ) 
	  //return ReportStgError( hr, L"CGostTable::Save", CLSID_CollTopics, IID_IGostTable );

   if( SUCCEEDED(hr) && fClearDirty == TRUE ) m_os = OS_None;

   return hr;
 }

STDMETHODIMP CGostTable::get_ClmName(BSTR *pVal) throw()
{
    return m_bsValClmName.CopyTo( pVal );
}

STDMETHODIMP CGostTable::put_ClmName(BSTR newVal) throw()
{
	if( m_bsValClmName == newVal ) return S_OK;

	m_bsValClmName = newVal;
	Modify();
	return S_OK;
}

STDMETHODIMP CGostTable::SetUniformGrad(VARIANT_BOOL Revers) throw()
{
	if( m_vec.size() < 1 )
	 {
	   //Error( L"Cann't assign values - the table is empty", IID_IGostTable, E_FAIL );
	   AtlReportError( CLSID_CollTopics, L"Cann't assign values - the table is empty", IID_IGostTable, E_FAIL );
	   return E_FAIL;
	 }

	if( m_vec.size() == 1 )
	  m_vec[ 0 ].m_fVal = 10;
	else
	 {
	   double dStep = 10.0 / (m_vec.size() - 1);
	   int i1, i2, iInc;

	   if( Revers == VARIANT_FALSE )
		 i1 = 0, i2 = m_vec.size(), iInc = 1;
	   else
		 i1 = m_vec.size() - 1, i2 = -1, iInc = -1;

	   for( int i = 0; i1 != i2; i1 += iInc, ++i )
		 m_vec[ i ].m_fVal = double(i1) * dStep;
	 }

	Modify();

	return S_OK;
}


STDMETHODIMP CGostTable::get_Storage( IStorage** Stg ) throw()
 {
   if( !Stg ) return E_POINTER;
   *Stg = NULL;
   return S_OK;
 }

STDMETHODIMP CGostTable::put_DefaultCU( /*[in]*/VARIANT_BOOL bMode ) throw()
 {   
   m_bDefaultCU = (bMode == VARIANT_TRUE ? true:false);
   return S_OK;
 }

STDMETHODIMP CGostTable::get_DefaultCU( /*[out,retval]*/VARIANT_BOOL* pbMode ) throw()
 {
   *pbMode = (m_bDefaultCU ? VARIANT_TRUE:VARIANT_FALSE);
   return S_OK;
 }

/*[local]*/BSTR STDMETHODCALLTYPE CGostTable::NameByRef( void )
 {
   return m_bsName;
 }

struct TFndByDescr: 
 public unary_function<VEC_CGostSlot::const_reference, bool>
  {
    TFndByDescr( VEC_CGostSlot::const_reference r ): m_ref( r )
	 {
	 }
    bool operator()( VEC_CGostSlot::const_reference r1 )
	 {
	   return (void*)&m_ref == (void*)&r1 ? false:
		 CmpBStrNoCase( m_ref.m_bsValDescr, r1.m_bsValDescr );
	 }

	VEC_CGostSlot::const_reference m_ref;
  };

STDMETHODIMP CGostTable::Check()
 {
   HRESULT hr = S_OK;
   m_shLastErr = -1;

   basic_stringstream<WCHAR> strm;
   
	   

   IT_VEC_CGostSlot it = m_vec.begin();
   for( int iCnt = 1; it != m_vec.end(); ++it, ++iCnt )
	 if( it->m_bsValDescr.m_str == NULL || it->m_bsValDescr.Length() < 1 )
	  {
	    m_shLastErr = iCnt - 1;
	    hr = E_FAIL;
		strm << L"Ошибка в строке " << iCnt << L". Пустое описание не допустимо (колонка 1)";
		AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, hr );
		break;
	  }
	 else if( it->m_bsLingvoVal.m_str == NULL || it->m_bsLingvoVal.Length() < 1 )
	  {
	    m_shLastErr = iCnt - 1;
	    hr = E_FAIL;
		strm << L"Ошибка в строке " << iCnt << L". Лингвистическое значение обязательно";
		AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, hr );
		break;
	  }
	 else
	  {
	    TFndByDescr dta( *it );
		IT_VEC_CGostSlot it2 = find_if( m_vec.begin(), m_vec.end(), dta ); 
		if( it2 != m_vec.end() )
		 {
		   m_shLastErr = iCnt - 1;
		   hr = E_FAIL;
		   strm << L"Дублируется описание \"" << Chk(it->m_bsValDescr) << L"\"";
		   AtlReportError( CLSID_CollTopics, strm.str().c_str(), IID_IGostTable, hr );
		   break;
		 }
	  }

   return hr;
 }


STDMETHODIMP CGostTable::get_LastErrStrZeroIdx(short *pVal)
{
	*pVal = m_shLastErr;
	return S_OK;
}
