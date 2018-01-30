// CompaundFilesSupport.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
//#include <comutil.h>
#include <comdef.h>

#include "CompaundFilesSupport.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}


typedef map<HRESULT, CComBSTR> TM_Cmp;

extern "C" COMPAUNDFILESSUPPORT_API void __fastcall GetStgErrDescr( HRESULT hr, LPCOLESTR wcPref, LPCOLESTR bsName, CComBSTR& rBs ) throw()
 {
   
   try {

	 const static TM_Cmp::value_type vtIni[ 17 ] = {
		 TM_Cmp::value_type( E_PENDING, L"Part or all of the necessary data is currently unavailable" ),
		 TM_Cmp::value_type( STG_E_ACCESSDENIED, L"Insufficient permissions to create storage object" ),	   
		 

		 TM_Cmp::value_type( STG_E_FILENOTFOUND, L"The specified file does not exist." ),
		 TM_Cmp::value_type( STG_E_LOCKVIOLATION, L"Access denied because another caller has the file open and locked"),
		 TM_Cmp::value_type( STG_E_SHAREVIOLATION ,L"Access denied because another caller has the file open and locked"),
		 TM_Cmp::value_type( STG_E_PATHNOTFOUND, L"Specified path does not exist"),
	   		 
		 
		 TM_Cmp::value_type( STG_E_FILEALREADYEXISTS, L"The name specified for the storage object already exists in the storage" ),
		 TM_Cmp::value_type( STG_E_INSUFFICIENTMEMORY, L"The storage object was not created due to a lack of memory."  ),
		 TM_Cmp::value_type( STG_E_INVALIDFLAG, L"The value specified for the grfMode flag is not a valid STGM enumeration value" ),
		 TM_Cmp::value_type( STG_E_INVALIDFUNCTION, L"The specified combination of grfMode flags is not supported." ),
		 TM_Cmp::value_type( STG_E_INVALIDNAME, L"Invalid value for pwcsName" ),
		 TM_Cmp::value_type( STG_E_INVALIDPOINTER, L"The pointer specified for the storage object was invalid" ),
		 TM_Cmp::value_type( STG_E_INVALIDPARAMETER, L"One of the parameters was invalid" ),
		 TM_Cmp::value_type( STG_E_REVERTED, L"The storage object has been invalidated by a revert operation above it in the transaction tree." ),
		 TM_Cmp::value_type( STG_E_TOOMANYOPENFILES, L"The storage object was not created because there are too many open files." ),

		 TM_Cmp::value_type( STG_E_MEDIUMFULL, L"There is insufficient disk space to complete operation." ),
		 TM_Cmp::value_type( STG_E_NOTCURRENT, L"The storage has been changed since the last commit." )
	  };

	 const static TM_Cmp mCmp( &vtIni[0], &vtIni[sizeof(vtIni)/sizeof(vtIni[0])] );

	 basic_stringstream<wchar_t> strm;

	 if( HRESULT_FACILITY(hr) != FACILITY_STORAGE )
	  {
		basic_string<TCHAR> cStr;
		WIN32ERR_TO_COMERR( cStr, hr & 0x0000FFFF );
		USES_CONVERSION;
		strm << T2W(cStr.c_str());
	  }
	 else
	  {
		TM_Cmp::const_iterator it = mCmp.find( hr );
		if( bsName )
		 {
		   if( it == mCmp.end() ) strm << wcPref << L": [" << (wchar_t*)bsName << L"] - " << L"Unknown error";
		   else strm << wcPref << L": [" << (wchar_t*)bsName << L"] - " << Chk(it->second);
		 }
		else
		 {
		   if( it == mCmp.end() ) strm << wcPref << L": Unknown error";
		   else strm << wcPref << L": " << Chk(it->second);
		 }
	  }

	 	 
	 rBs = strm.str().c_str();
	}
   catch( bad_alloc e )
	{
	  rBs = L"";
	}   
   catch( exception e )
	{	  
	  rBs = L"";
	}
 }

extern "C" COMPAUNDFILESSUPPORT_API HRESULT __fastcall WIN32ERR_TO_COMERR( std::basic_string<TCHAR>& rStr, DWORD dwErr ) throw()
 {   
   LPVOID lpMsgBuf;   
   DWORD dw = FormatMessage( 
	 FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
	 NULL,
	 dwErr,
	 MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
	 (LPTSTR) &lpMsgBuf,
	 0,
	 NULL 
   );

   rStr = dw ? (LPTSTR)lpMsgBuf:_T("<Cant't get error info>");
   //Error( dw ? T2W((LPTSTR)lpMsgBuf):L"", iid, hr );
   
   LocalFree( lpMsgBuf );

   return HRESULT_FROM_WIN32( dwErr );
 }

extern "C" COMPAUNDFILESSUPPORT_API HRESULT  __fastcall ReportStgError( HRESULT h, LPCOLESTR bsNameFunc, const CLSID& ridCls, const IID& iid, LPCOLESTR bsNameFile ) throw()
 {
   CComBSTR bsOut;
   GetStgErrDescr( h, bsNameFunc, bsNameFile == NULL ? L"":bsNameFile, bsOut );
   AtlReportError( ridCls, bsOut, iid, h );
   return h;
 }

extern "C" COMPAUNDFILESSUPPORT_API bool __fastcall ShortIntoVariant( VARIANT* pVar, const short sArg ) throw()
 {
   bool bRes = true;

   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_VT(pVar) == (VT_I2|VT_BYREF) ) 
		*V_I2REF(pVar) = sArg;

	  else if( V_VT(pVar) == (VT_I4|VT_BYREF) ) 
		*V_I4REF(pVar) = sArg;

	  else if( V_VT(pVar) == VT_I2 ) 
		V_I2(pVar) = sArg;

	  else if( V_VT(pVar) == VT_I4 ) 
		V_I4(pVar) = sArg;

	  else if( V_VT(pVar) == VT_EMPTY )
		V_VT(pVar) = VT_I2,
		V_I2(pVar) = sArg;

	  else bRes = false;	   
	}		   

   return bRes;
 }

extern "C" COMPAUNDFILESSUPPORT_API bool  __fastcall FloatIntoVariant( VARIANT* pVar, const float fArg ) throw()
 {
   bool bRes = true;

   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_VT(pVar) == (VT_R4|VT_BYREF) ) 
		*V_R4REF(pVar) = fArg;

	  else if( V_VT(pVar) == (VT_R8|VT_BYREF) ) 
		*V_R8REF(pVar) = fArg;

	  else if( V_VT(pVar) == VT_R4 ) 
		V_R4(pVar) = fArg;

	  else if( V_VT(pVar) == VT_R8 ) 
		V_R8(pVar) = fArg;

	  else if( V_VT(pVar) == VT_EMPTY )
		V_VT(pVar) = VT_R4,
		V_R4(pVar) = fArg;

	  else bRes = false;	   
	}		   

   return bRes;
 }

extern "C" COMPAUNDFILESSUPPORT_API HRESULT  __fastcall WSTRIntoVariant( VARIANT* pVar, LPCWSTR bsArg ) throw()
 {
   HRESULT hr = S_OK;
   
   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_VT(pVar) == (VT_BSTR|VT_BYREF) ) 
	   {
	     SysFreeString( *V_BSTRREF(pVar) );
		 
		 *V_BSTRREF(pVar) = SysAllocString( bsArg );
		 if( *V_BSTRREF(pVar) == NULL && bsArg != NULL ) hr = E_OUTOFMEMORY;		 
	   }	  	  
	  else if( V_VT(pVar) == VT_BSTR ) 
	   {
	     SysFreeString( V_BSTR(pVar) );
		 
		 V_BSTR(pVar) = SysAllocString( bsArg );
		 if( V_BSTR(pVar) == NULL && bsArg != NULL ) hr = E_OUTOFMEMORY;
		 
	   }

	  else if( V_VT(pVar) == VT_EMPTY )
	   {
		 V_VT(pVar) = VT_BSTR;
		 V_BSTR(pVar) = SysAllocString( bsArg );
		 if( V_BSTR(pVar) == NULL && bsArg != NULL ) hr = E_OUTOFMEMORY;
	   }

	  else hr = E_INVALIDARG;
	}		   
   else hr = S_FALSE;

   return hr;
 }

extern "C" COMPAUNDFILESSUPPORT_API BSTR*  __fastcall GetBSTRDst( VARIANT* pVar, HRESULT& hr ) throw()
 {
   hr = S_OK;
   BSTR* pbsRes = NULL;
   
   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_VT(pVar) == (VT_BSTR|VT_BYREF) ) 
	   {
	     SysFreeString( *V_BSTRREF(pVar) );		 
		 *V_BSTRREF(pVar) = NULL, pbsRes = V_BSTRREF(pVar);
	   }	  	  
	  else if( V_VT(pVar) == VT_BSTR ) 
	   {
	     SysFreeString( V_BSTR(pVar) );		 
		 V_BSTR(pVar) = NULL, pbsRes = &(V_BSTR(pVar));
	   }

	  else if( V_VT(pVar) == VT_EMPTY )
	   {
		 V_VT(pVar) = VT_BSTR;
		 V_BSTR(pVar) = NULL;
		 pbsRes = &(V_BSTR(pVar));
	   }

	  else hr = E_INVALIDARG;
	}		   
   else hr = S_FALSE;

   return pbsRes;
 }


extern "C" COMPAUNDFILESSUPPORT_API float* __fastcall GetFloatDst( VARIANT* pVar, HRESULT& hr ) throw()
 {
   hr = S_OK;
   float* fRes = NULL;
  
   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_VT(pVar) == (VT_R4|VT_BYREF) ) 
		fRes = V_R4REF(pVar);
	  
	  else if( V_VT(pVar) == VT_R4 )
		fRes = &(V_R4(pVar));	  

	  else if( V_VT(pVar) == VT_EMPTY )
		V_VT(pVar) = VT_R4,
		fRes = &(V_R4(pVar));

	  else hr = E_INVALIDARG;
	}		   
   else hr = S_FALSE;

   return fRes;   
 }

extern "C" COMPAUNDFILESSUPPORT_API double*  __fastcall GetDoubleDst( VARIANT* pVar, HRESULT& hr ) throw()
 {
   hr = S_OK;
   double* fRes = NULL;
  
   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_VT(pVar) == (VT_R8|VT_BYREF) ) 
		fRes = V_R8REF(pVar);
	  
	  else if( V_VT(pVar) == VT_R8 ) 
		fRes = &(V_R8(pVar));	  

	  else if( V_VT(pVar) == VT_EMPTY )
		V_VT(pVar) = VT_R8,
		fRes = &(V_R8(pVar));

	  else hr = E_INVALIDARG;
	}		   
   else hr = S_FALSE;

   return fRes;   
 }

extern "C" COMPAUNDFILESSUPPORT_API bool __fastcall IsEmpty( IStorage* pStm ) throw()
 {
   bool bRes = true;
   
   if( pStm != NULL )
	{
	  CComPtr<IEnumSTATSTG> spEnum;
	  HRESULT hr = pStm->EnumElements( 0, NULL, 0, &spEnum );
	  if( SUCCEEDED(hr) )	  
	   {	     
		 STATSTG stg; memset( &stg, 0, sizeof(STATSTG) );
		 while( spEnum->Next(1, &stg, NULL) == S_OK )
		  {		    
			if( stg.pwcsName ) CoTaskMemFree( stg.pwcsName ), stg.pwcsName = NULL;

			if( (stg.type & STGTY_STREAM) == STGTY_STREAM ||
			    (stg.type & STGTY_STORAGE) == STGTY_STORAGE
			  )
			 {
			   bRes = false;
			   break;
			 }
			 
			memset( &stg, 0, sizeof(STATSTG) );
		  }
	   }
	}

   return bRes;
 }


extern "C" COMPAUNDFILESSUPPORT_API HRESULT __fastcall EndTransaction( IStorage* pStg, HRESULT hrIn, STGC stgc ) throw()
 {
   HRESULT hr = S_OK;

   STATSTG stg;
   if( SUCCEEDED(pStg->Stat(&stg, STATFLAG_NONAME)) )
	{
	  if( (stg.grfMode&STGM_TRANSACTED) == STGM_TRANSACTED )
	   {
		 if( SUCCEEDED(hrIn) )
		   hr = pStg->Commit( stgc );
		 else
		   hr = pStg->Revert();		 
	   }
	}			
   return hr;
 }

extern "C" COMPAUNDFILESSUPPORT_API HRESULT __fastcall ObjExists( IStorage* pStg, DWORD dwKey, STGTY stgt ) throw()
 {
   HRESULT hr;

   _ASSERTE( pStg != NULL );

   wchar_t wcTmp[ 1024 ];
   _ultow( dwKey, wcTmp, 10 );

  if( stgt == STGTY_STORAGE )
   {
     CComPtr<IStorage> spStorage;
     hr = pStg->OpenStorage( wcTmp, NULL, 
	    STGM_READ|STGM_DIRECT|STGM_SHARE_DENY_NONE, NULL, 0, &spStorage );
	 spStorage.Release();
   }
  else
   {
     CComPtr<IStream> spStream;
     hr = pStg->OpenStream( wcTmp, 0,
		STGM_READ|STGM_DIRECT|STGM_SHARE_DENY_NONE, 0, &spStream );      
	 spStream.Release();
   }

   return SUCCEEDED(hr) ? S_OK:hr;
 }

extern "C" COMPAUNDFILESSUPPORT_API bool __fastcall LongFromVariant( const VARIANT* pVar, long& lArg ) throw()
 {
   bool bRes = true;

   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_ISBYREF(pVar) == VT_BYREF )
	    switch( V_VT(pVar) )
		 {
		   case VT_I4|VT_BYREF:
			 lArg = *V_I4REF(pVar);
			 break;
		   case VT_I2|VT_BYREF:
			 lArg = *V_I2REF(pVar);
			 break;
		   case VT_UI4|VT_BYREF:
			 lArg = *V_UI4REF(pVar);
			 break;
		   case VT_UI2|VT_BYREF:
			 lArg = *V_UI2REF(pVar);
			 break;
		   default:
			 bRes = false;
		 }
	  else
	    switch( V_VT(pVar) )
		 {
		   case VT_I4:
			 lArg = V_I4(pVar);
			 break;
		   case VT_I2:
			 lArg = V_I2(pVar);
			 break;
		   case VT_UI4:
			 lArg = V_UI4(pVar);
			 break;
		   case VT_UI2:
			 lArg = V_UI2(pVar);
			 break;
		   default:
			 bRes = false;
		 }	  
	}		   

   return bRes;
 }

extern "C" COMPAUNDFILESSUPPORT_API bool __fastcall VBoolFromVariant( const VARIANT* pVar, VARIANT_BOOL& vbArg ) throw()
 {
   bool bRes = true;

   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
	  if( V_VT(pVar) == (VT_BOOL|VT_BYREF) ) 
		vbArg = *V_BOOLREF(pVar);

	  else if( V_VT(pVar) == VT_BOOL ) 
		vbArg = V_BOOL(pVar);
	  	  
	  else bRes = false;	   
	}		   

   return bRes;
 }

extern "C" COMPAUNDFILESSUPPORT_API bool __fastcall FloatFromVariant( const VARIANT* pVar, float& lArg ) throw()
 {
   bool bRes = true;

   if( pVar != NULL && ((V_VT(pVar)&VT_TYPEMASK) != VT_ERROR) ) 
	{
      if( V_ISBYREF(pVar) == VT_BYREF )
	    switch( V_VT(pVar) )
		 {
		   case VT_R4|VT_BYREF:
			 lArg = *V_R4REF(pVar);
			 break;
		   case VT_R8|VT_BYREF:
			 lArg = *V_R8REF(pVar);
			 break;		   
		   default:
			 bRes = false;
		 }
	  else
	    switch( V_VT(pVar) )
		 {
		   case VT_R4:
			 lArg = V_R4(pVar);
			 break;
		   case VT_R8:
			 lArg = V_R8(pVar);
			 break;		   
		   default:
			 bRes = false;
		 }	  
	}		   

   return bRes;
 }

