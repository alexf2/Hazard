// CompaundFiles.cpp : Implementation of CCompaundFiles
#include "stdafx.h"
#include "GertNet.h"
#include "CompaundFiles.h"
#include "passerr.h"
#include "comutil.h"

#include <malloc.h>
#include <sstream>
#include <memory>
#include <list>
#include <map>
#include <algorithm>
using namespace std;



/////////////////////////////////////////////////////////////////////////////
// CCompaundFiles

STDMETHODIMP CCompaundFiles::InterfaceSupportsErrorInfo(REFIID riid)  throw()
{
	static const IID* arr[] = 
	{
		&IID_ICompaundFiles
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}



STDMETHODIMP CCompaundFiles::CreateCF(BSTR bsName, EnumCompConsts lAttrOpen, IStorage **ppIStorage)  throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !ppIStorage || !bsName ) return E_POINTER;

	
	HRESULT hh = StgCreateDocfile( bsName, 
	  lAttrOpen, NULL, ppIStorage );

	if( FAILED(hh) )
	   return ReportStgError( hh, L"CreateCF", GetObjectCLSID(), IID_ICompaundFiles, bsName );
	 /*{
	   CComBSTR bsTmp;
	   GetStgErrDescr( hh, L"CreateCF", bsName, bsTmp );	   	   
	   AtlReportError( GetObjectCLSID(), bsTmp, IID_ICompaundFiles, hh );
	   	   
	   return hh;
	 }*/

	return S_OK;
}

STDMETHODIMP CCompaundFiles::OpenCF(BSTR bsName, EnumCompConsts lAttrOpen, IStorage **ppIStorage)  throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !ppIStorage || !bsName ) return E_POINTER;

	
	HRESULT hh = StgOpenStorage( bsName, NULL, 
	  lAttrOpen, NULL, NULL, ppIStorage );

	if( FAILED(hh) )
	   return ReportStgError( hh, L"OpenCF", GetObjectCLSID(), IID_ICompaundFiles, bsName );	 

	return S_OK;
}

STDMETHODIMP CCompaundFiles::CreateStorageInt(/*[in]*/IStorage *pStg, /*[in]*/BSTR bsName, /*[in]*/EnumCompConsts lAttrOpen, /*[out, retval]*/IStorage** ppIStorage)  throw()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pStg || !ppIStorage || !bsName ) return E_POINTER;

	
	HRESULT hh = pStg->CreateStorage( bsName, lAttrOpen, 0, 0, ppIStorage );

	if( FAILED(hh) )
	  return ReportStgError( hh, L"CreateStorageInt", GetObjectCLSID(), IID_ICompaundFiles, bsName );	 
	 

	return S_OK;
 }

STDMETHODIMP CCompaundFiles::OpenStorageInt(/*[in]*/IStorage *pStg, /*[in]*/BSTR bsName, /*[in]*/EnumCompConsts lAttrOpen, /*[out, retval]*/IStorage** ppIStorage)  throw()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pStg || !ppIStorage || !bsName ) return E_POINTER;

	
	HRESULT hh = pStg->OpenStorage( bsName, NULL, lAttrOpen, 0, 0, ppIStorage );

	if( FAILED(hh) )
	  return ReportStgError( hh, L"OpenStorageInt", GetObjectCLSID(), IID_ICompaundFiles );	 
	 

	return S_OK;
 }

STDMETHODIMP CCompaundFiles::CreateStreamInt( IStorage *pStg, BSTR NameStream, EnumCompConsts lAttrOpen, IStream** ppIStrm ) throw()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	if( !pStg || !ppIStrm || !NameStream ) return E_POINTER;

	
	HRESULT hh = pStg->CreateStream( NameStream, lAttrOpen, 0, 0, ppIStrm );

	if( FAILED(hh) )
	  return ReportStgError( hh, L"CreateStreamInt", GetObjectCLSID(), IID_ICompaundFiles );
	 

	return S_OK;
 }

STDMETHODIMP CCompaundFiles::OpenStreamInt( IStorage *pStg, BSTR NameStream, EnumCompConsts lAttrOpen, IStream** ppIStrm ) throw()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	if( !pStg || !ppIStrm || !NameStream ) return E_POINTER;

	
	HRESULT hh = pStg->OpenStream( NameStream, NULL, lAttrOpen, 0, ppIStrm );

	if( FAILED(hh) )
	  return ReportStgError( hh, L"OpenStreamInt", GetObjectCLSID(), IID_ICompaundFiles );
	 

	return S_OK;
 }

static int __fastcall SzVar( VARIANT* pV )  throw()
 {      
   //switch( V_VT(pV) & ~VT_BYREF )
   switch( V_VT(pV) & VT_TYPEMASK )
	{
	  case VT_UI1: return sizeof(pV->bVal);
	  case VT_I2: return sizeof(pV->iVal);
	  case VT_I4: return sizeof(pV->lVal);
	  case VT_R4: return sizeof(pV->fltVal);
	  case VT_R8: return sizeof(pV->dblVal);
	  case VT_BOOL: return sizeof(pV->boolVal);
	  case VT_ERROR: return sizeof(pV->scode);
	  case VT_CY: return sizeof(pV->cyVal);
	  case VT_DATE: return sizeof(pV->date);
	  case VT_BSTR: return sizeof(pV->bstrVal);
	  case VT_UNKNOWN: return sizeof(pV->punkVal);
	  case VT_DISPATCH: return sizeof(pV->pdispVal);	  
	  default: return -1;
	}
 }

static void __fastcall VarExtractTo( VARIANT* pV, LPBYTE lpB ) throw()
 {
   void* pVoid = NULL;
   int iSz;
   /*if( (V_VT(pV) & VT_BYREF) == VT_BYREF )
	 pVoid = &pV->pbVal, iSz = sizeof(pV->pbVal);
   else*/
	 switch( V_VT(pV) )
	  {
		case VT_UI1: 
		  pVoid = &pV->bVal; iSz = sizeof(pV->bVal);
		  break;

		case VT_I2: 
		  pVoid = &pV->iVal; iSz = sizeof(pV->iVal);
		  break;
		case VT_I4: 
		  pVoid = &pV->lVal; iSz = sizeof(pV->lVal);
		  break;
		case VT_R4: 
		  pVoid = &pV->fltVal; iSz = sizeof(pV->fltVal);
		  break;
		case VT_R8: 
		  pVoid = &pV->dblVal; iSz = sizeof(pV->dblVal);
		  break;

		case VT_BOOL: 
		  pVoid = &pV->boolVal; iSz = sizeof(pV->boolVal);
		  break;

		case VT_ERROR: 
		  pVoid = &pV->scode; iSz = sizeof(pV->scode);
		  break;
		case VT_CY: 
		  pVoid = &pV->cyVal; iSz = sizeof(pV->cyVal);
		  break;
		case VT_DATE: 
		  pVoid = &pV->date; iSz = sizeof(pV->date);
		  break;
		case VT_BSTR: 
		  pVoid = &pV->bstrVal; iSz = sizeof(pV->bstrVal);
		  break;
		case VT_UNKNOWN: 
		  pVoid = &pV->punkVal; iSz = sizeof(pV->punkVal);
		  break;
		case VT_DISPATCH: 
		  pVoid = &pV->pdispVal; iSz = sizeof(pV->pdispVal);
		  break;		



		case VT_UI1|VT_BYREF: 
		  pVoid = pV->pbVal; iSz = sizeof(pV->bVal);
		  break;

		case VT_I2|VT_BYREF: 
		  pVoid = pV->piVal; iSz = sizeof(pV->iVal);
		  break;
		case VT_I4|VT_BYREF: 
		  pVoid = pV->plVal; iSz = sizeof(pV->lVal);
		  break;
		case VT_R4|VT_BYREF: 
		  pVoid = pV->pfltVal; iSz = sizeof(pV->fltVal);
		  break;
		case VT_R8|VT_BYREF: 
		  pVoid = pV->pdblVal; iSz = sizeof(pV->dblVal);
		  break;

		case VT_BOOL|VT_BYREF: 
		  pVoid = pV->pboolVal; iSz = sizeof(pV->boolVal);
		  break;

		case VT_ERROR|VT_BYREF: 
		  pVoid = pV->pscode; iSz = sizeof(pV->scode);
		  break;
		case VT_CY|VT_BYREF: 
		  pVoid = pV->pcyVal; iSz = sizeof(pV->cyVal);
		  break;
		case VT_DATE|VT_BYREF: 
		  pVoid = pV->pdate; iSz = sizeof(pV->date);
		  break;
		case VT_BSTR|VT_BYREF: 
		  pVoid = pV->pbstrVal; iSz = sizeof(pV->bstrVal);
		  break;
		case VT_UNKNOWN|VT_BYREF: 
		  pVoid = pV->ppunkVal; iSz = sizeof(pV->punkVal);
		  break;
		case VT_DISPATCH|VT_BYREF: 
		  pVoid = pV->ppdispVal; iSz = sizeof(pV->pdispVal);
		  break;		
	  }

   int iSzAln = (iSz + sizeof(int) - 1) & ~(sizeof(int) - 1);
   memset( lpB, 0, iSzAln );
   memcpy( lpB /*+(iSzAln - iSz)*/, pVoid, iSz );
 }

STDMETHODIMP CCompaundFiles::Sprintf(BSTR bsFormatString, long lBuffSize, SAFEARRAY** pparrArgs, BSTR* bsResult) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	HRESULT hr = S_OK;  
	
	if( !bsFormatString || !pparrArgs || !bsResult ) hr = E_POINTER;	
	else
     {
	   hr = SafeArrayLock( *pparrArgs );
	   if( FAILED(hr) )
		{	  
		  Error( L"Cann't array access", IID_IMGertNet, E_FAIL );
		  hr = E_FAIL;
		}
	   else
		{
		  if( (*pparrArgs)->cDims != 1 )
		   {
			 SafeArrayUnlock( *pparrArgs );
			 Error( L"Array must have only 1 dimension", IID_IMGertNet, E_FAIL );
			 hr = E_FAIL;
		   }
		  else
		   {
			 if( ((*pparrArgs)->fFeatures&FADF_VARIANT) != FADF_VARIANT )
			  {
				SafeArrayUnlock( *pparrArgs );
				Error( L"Bad array type: Array must be a VARIANTs array", IID_IMGertNet, E_FAIL );
				hr = E_FAIL;
			  }
			 else
			  { 
				VARIANT* pV = (VARIANT*)(*pparrArgs)->pvData; 
				int iSz = 0;
				for( int i = (*pparrArgs)->rgsabound[0].cElements; i > 0; --i, ++pV )
				 {
				   int iSzVar = SzVar( pV );
				   if( iSzVar == -1 )
					{
					  SafeArrayUnlock( *pparrArgs );
					  Error( L"In var-parameters unsupported type", IID_IMGertNet, E_FAIL );
					  hr = E_FAIL;
					  break;
					}
				   iSz += (iSzVar + sizeof(int) - 1) & ~(sizeof(int) - 1);
				 }
				
				if( SUCCEEDED(hr) )
				 {	   				   
				   //auto_ptr<BYTE> apDta( new BYTE[iSz] );
				   LPBYTE pDtaPrms;
				   try {
					  pDtaPrms = (LPBYTE)_alloca( iSz * sizeof(BYTE) );
					}
				   catch( ... )
					{
					  //hr = E_OUTOFMEMORY;		  
					  pDtaPrms = 0;
					}

				   if( pDtaPrms == NULL ) hr = E_OUTOFMEMORY;
                   else 				   
					{
					  //LPBYTE lpB = apDta.get();
					  LPBYTE lpB = pDtaPrms;
					  pV = (VARIANT*)(*pparrArgs)->pvData; 	
					  for( i = (*pparrArgs)->rgsabound[0].cElements; i > 0; --i, ++pV )
					   {
						 VarExtractTo( pV, lpB );
						 lpB += (SzVar(pV) + sizeof(int) - 1) & ~(sizeof(int) - 1);
					   }

					  SafeArrayUnlock( *pparrArgs );

					  //va_list vLst = (va_list)apDta.get();
					  va_list vLst = (va_list)pDtaPrms;
					  //int iSizeRet = 1024L * 10L;
					  //auto_ptr<wchar_t> apTmpStr( new wchar_t[lBuffSize] );
					  LPWSTR  lpwTmpStr;
					  try {
						 lpwTmpStr = (LPWSTR)_alloca( lBuffSize * sizeof(wchar_t) );
					   }
					  catch( ... )
					   {						 
						 lpwTmpStr = NULL;
					   }

					  if( lpwTmpStr == NULL ) hr = E_OUTOFMEMORY;		   
                      else					  
					   {
						 int iSzWritten =
						   _vsnwprintf( lpwTmpStr, lBuffSize - 1, bsFormatString, vLst );
						 if( iSzWritten < 0 )
						  {
							Error( L"Error from _vsnwprintf: probably buffer's size is not enought", IID_IMGertNet, E_FAIL );
							hr = E_FAIL;
						  }
						 else
						  {						 
							*bsResult = SysAllocStringLen( lpwTmpStr, iSzWritten );	
 
							if( *bsResult == NULL && iSzWritten > 0 )
							 {
							   Error( L"No memory for result", IID_IMGertNet, E_OUTOFMEMORY );
							   hr = E_OUTOFMEMORY;
							 }
						  }
					   }
					}
				 }
			  }
		   }
		}
	 }
	 

	return hr;
}

STDMETHODIMP CCompaundFiles::HiLong(long lArg, long *pRes) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pRes ) return E_POINTER;
	*pRes = HIWORD(lArg);

	return S_OK;
}

STDMETHODIMP CCompaundFiles::LoLong(long lArg, long *pRes) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pRes ) return E_POINTER;
	*pRes = LOWORD(lArg);

	return S_OK;
}


STDMETHODIMP CCompaundFiles::LShiftLong(long lArg, long lCnt, long *pRes) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pRes ) return E_POINTER;
	*(unsigned long*)pRes = ((unsigned long)lArg) << (unsigned long)lCnt;

	return S_OK;
}

STDMETHODIMP CCompaundFiles::RShiftLong(long lArg, long lCnt, long *pRes) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	if( !pRes ) return E_POINTER;
	*(unsigned long*)pRes = ((unsigned long)lArg) >> (unsigned long)lCnt;

	return S_OK;
}

static HRESULT CreArrrBSTR( CCompaundFiles* pCls, list<BSTR>& rL, SAFEARRAY** ppArr ) throw()
 {

   HRESULT hr = S_OK;

   SAFEARRAYBOUND bnd = { rL.size(), 0 };
   *ppArr = SafeArrayCreate( VT_BSTR, 1, &bnd );

   if( !*ppArr )
	{
	  pCls->Error( L"Cann't array allocate", IID_ICompaundFiles, E_FAIL );
	  hr = E_FAIL;
	}
   else
	{
	  BSTR* pDta;
	  hr = SafeArrayAccessData( *ppArr, (void**)&pDta );
	  if( FAILED(hr) )
	   {
		 SafeArrayDestroy( *ppArr );
		 *ppArr = NULL;	   
		 pCls->Error( L"Cann't array access", IID_ICompaundFiles, hr );		 
	   }
	  else
	   {
		 copy( rL.begin(), rL.end(), pDta );
		 SafeArrayUnaccessData( *ppArr );
	   }
	}

   return hr;
 }

static void SFree( BSTR bs ) throw()
 {
   SysFreeString( bs );
 }

STDMETHODIMP CCompaundFiles::DirOfStorage(IStorage *pStg, TDRFlags shFlags, SAFEARRAY** ppArrNamesStg, SAFEARRAY** ppArrNamesStrm) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

    HRESULT hr = S_OK;

	list<BSTR> lstStor;
	list<BSTR> lstStrm;

	try {
	   if( !ppArrNamesStrm && !ppArrNamesStg ) return E_POINTER;
	   

	   CComPtr<IEnumSTATSTG> spEnum;
	   hr = pStg->EnumElements( 0, 0, 0, &spEnum );
	   if( SUCCEEDED(hr) )
		{
		  STATSTG stat;
		  while( spEnum->Next(1, &stat, 0) == S_OK )
		   {
			 if( (shFlags&TDR_Stg) && stat.type == STGTY_STORAGE ) 
			  {
				BSTR bsTmp_ = SysAllocString( stat.pwcsName );
				if( bsTmp_ == NULL )
				 {
				   CoTaskMemFree( stat.pwcsName );
				   hr = E_OUTOFMEMORY;
				   break;
				 }
				else
				  lstStor.push_back( bsTmp_ ); 
			  }
			 else if( (shFlags&TDR_Strm) && stat.type == STGTY_STREAM ) 
			  {
				BSTR bsTmp_ = SysAllocString( stat.pwcsName );
				if( bsTmp_ == NULL )
				 {
				   CoTaskMemFree( stat.pwcsName );
				   hr = E_OUTOFMEMORY;
				   break;
				 }
				else
				  lstStrm.push_back( bsTmp_ ); 
			  }

			 if( stat.pwcsName ) CoTaskMemFree( stat.pwcsName );
		   }
		}

	   if( FAILED(hr) ) 
		 ReportStgError( hr, L"DirOfStorage", GetObjectCLSID(), IID_ICompaundFiles );	 	 
	   else
		{
		  if( shFlags&TDR_Stg )
		   {
			 hr = CreArrrBSTR( this, lstStor, ppArrNamesStg );
			 if( FAILED(hr) )
				for_each( lstStor.begin(), lstStor.end(), SFree );
		   }		  

		  if( shFlags&TDR_Strm )
		   {
			 hr = CreArrrBSTR( this, lstStrm, ppArrNamesStrm );
			 if( FAILED(hr) )
				for_each( lstStrm.begin(), lstStrm.end(), SFree );
		   }
		}
	 }
	catch( bad_alloc e )
	 {
	   hr = E_OUTOFMEMORY;
	   for_each( lstStor.begin(), lstStor.end(), SFree );
	   for_each( lstStrm.begin(), lstStrm.end(), SFree );
	 }
	catch( exception e )
	{	  
	  USES_CONVERSION;
	  hr = E_FAIL;
	  AtlReportError( GetObjectCLSID(), A2COLE(e.what()), IID_ICompaundFiles, hr );
	}
	catch( CMemoryException* pE )
	 {
	   hr = E_OUTOFMEMORY;
	   pE->Delete();
	   for_each( lstStor.begin(), lstStor.end(), SFree );
	   for_each( lstStrm.begin(), lstStrm.end(), SFree );
	 }
	 
	return hr;
}

STDMETHODIMP CCompaundFiles::Win32ErrInfo(long ErrCode,  BSTR *ResultStr) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

    USES_CONVERSION; 
    if( ResultStr == NULL ) return E_POINTER;

    std::basic_string<TCHAR> sTmp;
	WIN32ERR_TO_COMERR( sTmp, ErrCode );

    *ResultStr = SysAllocString( A2W(sTmp.c_str()) );

	return *ResultStr == NULL ? E_OUTOFMEMORY:S_OK;
}

STDMETHODIMP CCompaundFiles::CutPath(BSTR FullPath, VARIANT* Dir, VARIANT* Name, VARIANT* Ext)
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


    if( FullPath == NULL ) return E_POINTER;
	wchar_t wcDrive[_MAX_DRIVE], wcDir[_MAX_DIR], wcName[_MAX_FNAME], wcExt[_MAX_EXT];
	wcDrive[0] = wcDir[0] = wcName[0] = wcExt[0] = 0;
	_wsplitpath( FullPath, wcDrive, wcDir, wcName, wcExt );


	HRESULT hr = S_OK;

	hr = WSTRIntoVariant( Ext, wcExt );
	if( hr == E_INVALIDARG )
	  Error( L"Bad parametr 'Ext' type. String is expected", IID_ICompaundFiles, E_INVALIDARG );
	else if( SUCCEEDED(hr) )
	 {
	   hr = WSTRIntoVariant( Name, wcName );
	   if( hr == E_INVALIDARG )
		 Error( L"Bad parametr 'Name' type. String is expected", IID_ICompaundFiles, E_INVALIDARG );
	   else if( SUCCEEDED(hr) )
		{
		  hr = WSTRIntoVariant( Dir, L"" );
		  if( hr == E_INVALIDARG )
		    Error( L"Bad parametr 'Dir' type. String is expected", IID_ICompaundFiles, E_INVALIDARG );
		  else if( SUCCEEDED(hr) && hr != S_FALSE )
		   {
		     LPWSTR lpw;
			 try {
                lpw = (LPWSTR)_alloca( 2 * (_MAX_DRIVE + _MAX_DIR + 1) );
			  }
			 catch( ... )
			  {
			    lpw = 0;
			  }
			 
			 if( lpw == NULL ) hr = E_OUTOFMEMORY;
			 else
			  {
				*lpw = 0;
				wcscpy( lpw, wcDrive );
				wcscat( lpw, wcDir );
				hr = WSTRIntoVariant( Dir, lpw );
			  }
		   }
		}
	 }
	 
			
	return SUCCEEDED(hr) ? S_OK:hr;
}

STDMETHODIMP CCompaundFiles::Commit(STGC1 Flags, IStorage *Stg) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	HRESULT hr = S_OK;
	if( Stg == NULL ) 
	  hr = E_POINTER;
	else
	 {
	   hr = Stg->Commit( Flags );
	   if( FAILED(hr) )
		 ReportStgError( hr, L"Commit", GetObjectCLSID(), IID_ICompaundFiles );	 	 ReportStgError( hr, L"DirOfStorage", GetObjectCLSID(), IID_ICompaundFiles );	 	 
	 }

	return hr;
}

STDMETHODIMP CCompaundFiles::Revert(IStorage *Stg) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif


	HRESULT hr = S_OK;
	if( Stg == NULL ) 
	  hr = E_POINTER;
	else
	 {
	   hr = Stg->Revert();
	   if( FAILED(hr) )
		 ReportStgError( hr, L"Revert", GetObjectCLSID(), IID_ICompaundFiles );	 	 ReportStgError( hr, L"DirOfStorage", GetObjectCLSID(), IID_ICompaundFiles );	 	 
	 }

	return hr;
}

STDMETHODIMP CCompaundFiles::PutString( /*[in]*/IStream* Strm, /*[in]*/BSTR Str ) throw()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif

	HRESULT hr = S_OK;
	if( Strm == NULL || Str == NULL ) 
	  hr = E_POINTER;
	else
	 {
	   CComBSTR bsStr; bsStr.Attach( Str );
	   hr = bsStr.WriteToStream( Strm );
	   if( FAILED(hr) )
		 ReportStgError( hr, L"PutString", GetObjectCLSID(), IID_ICompaundFiles );	 	 ReportStgError( hr, L"DirOfStorage", GetObjectCLSID(), IID_ICompaundFiles );

	   bsStr.Detach();
	 }

	return hr;
 }

STDMETHODIMP CCompaundFiles::GetString( /*[in]*/IStream* Strm, /*[out,retval]*/BSTR* pStr ) throw()
 {
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif  

	HRESULT hr = S_OK;
	if( Strm == NULL || pStr == NULL ) 
	  hr = E_POINTER;
	else
	 {
	   CComBSTR bsStr; 
	   hr = bsStr.ReadFromStream( Strm );
	   if( FAILED(hr) )
		{
		  ReportStgError( hr, L"PutString", GetObjectCLSID(), IID_ICompaundFiles );	 	 ReportStgError( hr, L"DirOfStorage", GetObjectCLSID(), IID_ICompaundFiles );
		  *pStr = NULL;
		}
	   else
	     *pStr = bsStr.Detach();
	 }

	return hr;
 }


STDMETHODIMP CCompaundFiles::DestroyCF( IStorage* Stg, BSTR Name ) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif  

    if( !Stg || !Name ) return E_POINTER;

	
	HRESULT hh = Stg->DestroyElement( Name );

	if( FAILED(hh) )
	  return ReportStgError( hh, L"DestroyCF", GetObjectCLSID(), IID_ICompaundFiles, Name );
	 
	return S_OK;
}

STDMETHODIMP CCompaundFiles::IsEmptyStg( IStorage *Stg, VARIANT_BOOL* bEmpty ) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif  

	if( !Stg || !bEmpty ) return E_POINTER;

	*bEmpty = (IsEmpty(Stg) ? VARIANT_TRUE:VARIANT_FALSE);

	return S_OK;
}

STDMETHODIMP CCompaundFiles::MakeLong(long Lo, long Hi, long *lRes) throw()
{
#ifdef _AFXDLL
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
#endif  


	*(unsigned long*)lRes = MAKELONG( Lo, Hi );

	return S_OK;
}
