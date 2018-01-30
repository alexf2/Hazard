#include "stdafx.h"
#include "fi_private.h"


using namespace FI_API;

TInitialize tiInit;
LPTSTR lpCurrFile = NULL;

TFntDescr tfdFonts[] =
 {
   { _T("Tahoma"), _T("Tahoma"), _T(".ttf"), _T("Tahoma (True Type)"),
	 FNT_TAHOMA, RUSSIAN_CHARSET, DEFAULT_PITCH|FF_DONTCARE, TRUETYPE_FONTTYPE , true }
 };

size_t sz_tfdFonts = sizeof(tfdFonts) / sizeof(tfdFonts[0]);

void __fastcall GetErrInfo( std::basic_string<_TCHAR>& rStr, DWORD dwErr )
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

   rStr = dw ? (LPTSTR)lpMsgBuf:_T("<Ошибка извлечения информации об ошибке>");   
   
   LocalFree( lpMsgBuf );   
 }

LONG __fastcall ErrRpt( FIAPI_Errors err ) 
 {   
   switch( err )
	{
	  case FIAPI_NoMemory:
	    tiInit.SetError( _T("Не хватает памяти") ); 
	    break;
	  
	  default:
	   {
		 DWORD dw = GetLastError();		 		 
		 std::basic_string<TCHAR> str;
		 GetErrInfo( str, dw );
		 
		 if( err == FIAPI_CantUpdateRegistry )
		  {
		    std::basic_stringstream<TCHAR> strm;
			strm << _T("Ошибка записи в реестр (Fonts): ") << str;
		    tiInit.SetError( strm.str().c_str() );
		  }
		 else if( err == FIAPI_IOError && lpCurrFile != NULL )
		  {
		    std::basic_stringstream<TCHAR> strm;
			strm << _T("Ошибка при модификации: '") << lpCurrFile << _T("' - ") << str;
		    tiInit.SetError( strm.str().c_str() );
			lpCurrFile = NULL;
		  }
		 else   
		   tiInit.SetError( str.c_str() );
	   }
	 }

   return err;
  }
